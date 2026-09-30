use super::{
    Context, ExportArgs, InsertArgs, Job, OpenArgs, Path, PathArgs, ReadArgs, ReplaceArgs, Result,
    UndoArgs, Value, bail, document_path, json, mpsc, output_path, parse, word_find_text,
};
use ::windows::{
    Win32::{
        Foundation::VARIANT_BOOL,
        System::{
            Com::{
                CLSCTX_LOCAL_SERVER, CLSIDFromProgID, COINIT_APARTMENTTHREADED, CoCreateInstance,
                CoInitializeEx, CoUninitialize, DISPATCH_FLAGS, DISPATCH_METHOD,
                DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO, IDispatch,
            },
            Ole::GetActiveObject,
            Variant::{VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_I2, VT_I4, VT_INT, VariantClear},
        },
        UI::WindowsAndMessaging::{
            DispatchMessageW, MSG, PM_REMOVE, PeekMessageW, TranslateMessage,
        },
    },
    core::{BSTR, GUID, Interface, PCWSTR, w},
};
use cppvtable_com::ComPtr;
use std::{ffi::c_void, marker::PhantomData, mem::ManuallyDrop, rc::Rc, time::Duration};

#[cfg(test)]
mod docx_compat;

// VARIANT does not implement Drop in windows-rs. This owner clears every BSTR/interface.
#[repr(transparent)]
struct Variant(VARIANT);
impl Variant {
    fn empty() -> Self {
        Self(VARIANT::default())
    }
    fn int(n: i32) -> Self {
        let mut v = Self::empty();
        unsafe {
            (*v.0.Anonymous.Anonymous).vt = VT_I4;
            (*v.0.Anonymous.Anonymous).Anonymous.lVal = n;
        }
        v
    }
    fn boolean(b: bool) -> Self {
        let mut v = Self::empty();
        unsafe {
            (*v.0.Anonymous.Anonymous).vt = VT_BOOL;
            (*v.0.Anonymous.Anonymous).Anonymous.boolVal = VARIANT_BOOL(if b { -1 } else { 0 });
        }
        v
    }
    fn text(s: &str) -> Self {
        let mut v = Self::empty();
        unsafe {
            (*v.0.Anonymous.Anonymous).vt = VT_BSTR;
            (*v.0.Anonymous.Anonymous).Anonymous.bstrVal = ManuallyDrop::new(BSTR::from(s));
        }
        v
    }
    fn integer(&self) -> Result<i32> {
        unsafe {
            let a = &self.0.Anonymous.Anonymous;
            match a.vt {
                VT_I4 | VT_INT => Ok(a.Anonymous.lVal),
                VT_I2 => Ok(i32::from(a.Anonymous.iVal)),
                _ => bail!("Word returned a noninteger value"),
            }
        }
    }
    fn bool_value(&self) -> Result<bool> {
        unsafe {
            let a = &self.0.Anonymous.Anonymous;
            if a.vt != VT_BOOL {
                bail!("Word returned a nonboolean value");
            }
            Ok(a.Anonymous.boolVal.0 != 0)
        }
    }
    fn string(&self) -> Result<String> {
        unsafe {
            let a = &self.0.Anonymous.Anonymous;
            if a.vt != VT_BSTR {
                bail!("Word returned a nontext value");
            }
            Ok(a.Anonymous.bstrVal.to_string())
        }
    }
    fn dispatch(&self) -> Result<Dispatch> {
        unsafe {
            let a = &self.0.Anonymous.Anonymous;
            if a.vt != VT_DISPATCH {
                bail!("Word returned a nonobject value");
            }
            Dispatch::from_borrowed(
                a.Anonymous
                    .pdispVal
                    .as_ref()
                    .context("Word returned a null object")?
                    .as_raw(),
            )
        }
    }
}
impl Drop for Variant {
    fn drop(&mut self) {
        unsafe {
            let _ = VariantClear(&raw mut self.0);
        }
    }
}

// The four IDispatch slots follow the inherited IUnknown slots in the COM ABI. Every
// method takes raw pointers whose validity the caller must guarantee, so all are unsafe.
#[cppvtable_com::interface(abi = com, iid = "00020400-0000-0000-c000-000000000046")]
unsafe trait IAutomationDispatch {
    unsafe fn get_type_info_count(&self, count: *mut u32) -> ::windows::core::HRESULT;
    unsafe fn get_type_info(
        &self,
        index: u32,
        locale: u32,
        info: *mut *mut c_void,
    ) -> ::windows::core::HRESULT;
    unsafe fn get_ids_of_names(
        &self,
        iid: *const GUID,
        names: *const PCWSTR,
        count: u32,
        locale: u32,
        ids: *mut i32,
    ) -> ::windows::core::HRESULT;
    unsafe fn invoke(
        &self,
        id: i32,
        iid: *const GUID,
        locale: u32,
        flags: u16,
        params: *const DISPPARAMS,
        result: *mut VARIANT,
        exception: *mut EXCEPINFO,
        argument: *mut u32,
    ) -> ::windows::core::HRESULT;
}

#[derive(Clone)]
struct Dispatch(ComPtr<IAutomationDispatch>, PhantomData<Rc<()>>);
impl Dispatch {
    fn from_windows(pointer: IDispatch) -> Self {
        // Transfer the OS boundary's owned IDispatch reference into cppvtable RAII.
        Self(
            unsafe { ComPtr::from_raw_unchecked(pointer.into_raw()) },
            PhantomData,
        )
    }
    unsafe fn from_borrowed(pointer: *mut c_void) -> Result<Self> {
        // A result VARIANT keeps the borrowed pointer alive during AddRef.
        Ok(Self(
            unsafe { ComPtr::from_raw_add_ref(pointer) }.context("Null Word dispatch pointer")?,
            PhantomData,
        ))
    }
    fn ids(&self, names: &[&str]) -> Result<Vec<i32>> {
        let wide: Vec<Vec<u16>> = names
            .iter()
            .map(|n| n.encode_utf16().chain(Some(0)).collect())
            .collect();
        let pointers: Vec<PCWSTR> = wide.iter().map(|s| PCWSTR(s.as_ptr())).collect();
        let mut ids = vec![0; names.len()];
        unsafe {
            self.0
                .get_ids_of_names(
                    &GUID::zeroed(),
                    pointers.as_ptr(),
                    u32::try_from(pointers.len())?,
                    0x409,
                    ids.as_mut_ptr(),
                )
                .ok()
        }
        .with_context(|| format!("Word member {} unavailable", names[0]))?;
        Ok(ids)
    }
    fn invoke(
        &self,
        name: &str,
        flags: DISPATCH_FLAGS,
        mut args: Vec<Variant>,
        mut named: Vec<i32>,
        id: i32,
    ) -> Result<Variant> {
        args.reverse();
        named.reverse();
        let params = DISPPARAMS {
            rgvarg: args.as_mut_ptr().cast(),
            rgdispidNamedArgs: named.as_mut_ptr(),
            cArgs: u32::try_from(args.len())?,
            cNamedArgs: u32::try_from(named.len())?,
        };
        // Retry only explicit busy rejections, which indicate the call was not executed.
        for attempt in 0..50 {
            let mut result = Variant::empty();
            let mut exception = EXCEPINFO::default();
            let mut argument = 0;
            let status = unsafe {
                self.0
                    .invoke(
                        id,
                        &GUID::zeroed(),
                        0x409,
                        flags.0,
                        &raw const params,
                        &raw mut result.0,
                        &raw mut exception,
                        &raw mut argument,
                    )
                    .ok()
            };
            let description = unsafe {
                if let Some(fill) = exception.pfnDeferredFillIn {
                    let _ = fill(&raw mut exception);
                }
                let description = exception.bstrDescription.to_string();
                ManuallyDrop::drop(&mut exception.bstrSource);
                ManuallyDrop::drop(&mut exception.bstrDescription);
                ManuallyDrop::drop(&mut exception.bstrHelpFile);
                description
            };
            match status {
                Ok(()) => return Ok(result),
                Err(e)
                    if matches!(e.code().0.cast_unsigned(), 0x8001_0001 | 0x8001_010a)
                        && attempt < 49 =>
                {
                    pump();
                    std::thread::sleep(Duration::from_millis(100));
                }
                Err(e) => bail!(
                    "Word {name} failed: {e}; {description} (argument {argument}). Dismiss any open Word dialog and retry."
                ),
            }
        }
        unreachable!()
    }
    fn get(&self, name: &str) -> Result<Variant> {
        self.invoke(
            name,
            DISPATCH_PROPERTYGET,
            vec![],
            vec![],
            self.ids(&[name])?[0],
        )
    }
    fn put(&self, name: &str, value: Variant) -> Result<()> {
        self.invoke(
            name,
            DISPATCH_PROPERTYPUT,
            vec![value],
            vec![-3],
            self.ids(&[name])?[0],
        )?;
        Ok(())
    }
    fn method(&self, name: &str, args: Vec<Variant>) -> Result<Variant> {
        self.invoke(name, DISPATCH_METHOD, args, vec![], self.ids(&[name])?[0])
    }
    fn named_method(&self, name: &str, names: &[&str], args: Vec<Variant>) -> Result<Variant> {
        let mut all = vec![name];
        all.extend_from_slice(names);
        let ids = self.ids(&all)?;
        self.invoke(name, DISPATCH_METHOD, args, ids[1..].to_vec(), ids[0])
    }
    fn object(&self, name: &str) -> Result<Self> {
        self.get(name)?.dispatch()
    }
}

fn pump() {
    unsafe {
        let mut msg = MSG::default();
        while PeekMessageW(&raw mut msg, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&raw const msg);
            DispatchMessageW(&raw const msg);
        }
    }
}
struct Apartment;
impl Drop for Apartment {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

pub(super) fn run(rx: mpsc::Receiver<Job>) {
    let initialized = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }.ok();
    if let Err(error) = initialized {
        for job in rx {
            let _ = job.reply.send(Err(format!(
                "Cannot initialize Word COM apartment: {error}"
            )));
        }
        return;
    }
    let _apartment = Apartment;
    let mut state = State { app: None };
    loop {
        pump();
        match rx.recv_timeout(Duration::from_millis(25)) {
            Ok(job) => {
                let result = state
                    .call(&job.name, &job.args)
                    .map_err(|e| format!("{e:#}"));
                let _ = job.reply.send(result);
            }
            Err(mpsc::RecvTimeoutError::Timeout) => {}
            Err(mpsc::RecvTimeoutError::Disconnected) => break,
        }
    }
    // State (and every COM pointer) is dropped before CoUninitialize. Never Quit Word.
}

struct State {
    app: Option<Dispatch>,
}
#[derive(Debug)]
struct WordNotRunning;
impl std::fmt::Display for WordNotRunning {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str("Microsoft Word is not running. Call word_live_open first.")
    }
}
impl std::error::Error for WordNotRunning {}
impl State {
    fn attach(&mut self, launch: bool) -> Result<Dispatch> {
        if let Some(app) = &self.app {
            if app.get("Version").is_ok() {
                return Ok(app.clone());
            }
            self.app = None;
        }
        let clsid = unsafe { CLSIDFromProgID(w!("Word.Application")) }
            .context("Desktop Microsoft Word is not installed/registered")?;
        let mut unknown = None;
        let active = unsafe { GetActiveObject(&raw const clsid, None, &raw mut unknown) };
        let app = match active {
            Ok(()) => Dispatch::from_windows(
                unknown
                    .context("Running Word did not expose automation")?
                    .cast::<IDispatch>()?,
            ),
            Err(e) if e.code().0.cast_unsigned() == 0x8004_01e3 && launch => {
                Dispatch::from_windows(
                    unsafe {
                        CoCreateInstance::<_, IDispatch>(
                            &raw const clsid,
                            None,
                            CLSCTX_LOCAL_SERVER,
                        )
                    }
                    .context("Cannot launch Microsoft Word")?,
                )
            }
            Err(e) if e.code().0.cast_unsigned() == 0x8004_01e3 => {
                return Err(WordNotRunning.into());
            }
            Err(e) => return Err(e).context("Cannot attach running Microsoft Word"),
        };
        self.app = Some(app.clone());
        Ok(app)
    }
    fn document(&mut self, path: &str) -> Result<(Dispatch, Dispatch)> {
        let target = document_path(path)?;
        let app = self.attach(false)?;
        let documents = app.object("Documents")?;
        let count = documents.get("Count")?.integer()?;
        let mut found = None;
        for index in 1..=count {
            let doc = documents
                .method("Item", vec![Variant::int(index)])?
                .dispatch()?;
            let full = doc.get("FullName")?.string()?;
            if Path::new(&full).is_absolute() && same_path(&target, Path::new(&full)) {
                if found.is_some() {
                    bail!(
                        "Multiple open documents match this path; close the duplicate before editing"
                    );
                }
                found = Some(doc);
            }
        }
        Ok((
            app,
            found.context(
                "The exact document is not open in Word; call word_live_open with its path",
            )?,
        ))
    }
    #[expect(
        clippy::too_many_lines,
        reason = "The explicit tool router keeps COM lifetimes and cleanup local to each operation"
    )]
    fn call(&mut self, name: &str, args: &Value) -> Result<Value> {
        match name {
            "word_live_status" => {
                let installed = unsafe { CLSIDFromProgID(w!("Word.Application")) }.is_ok();
                if !installed {
                    return Ok(json!({"installed":false,"running":false,"documents":[]}));
                }
                let app = match self.attach(false) {
                    Ok(app) => app,
                    Err(e) if e.is::<WordNotRunning>() => {
                        return Ok(json!({"installed":true,"running":false,"documents":[]}));
                    }
                    Err(e) => return Err(e),
                };
                let docs = app.object("Documents")?;
                let mut list = Vec::new();
                for index in 1..=docs.get("Count")?.integer()? {
                    let d = docs.method("Item", vec![Variant::int(index)])?.dispatch()?;
                    list.push(json!({"name":d.get("Name")?.string()?,"path":d.get("FullName")?.string()?,"saved":d.get("Saved")?.bool_value()?,"read_only":d.get("ReadOnly")?.bool_value()?}));
                }
                Ok(
                    json!({"installed":installed,"running":true,"version":app.get("Version")?.string()?,"documents":list}),
                )
            }
            "word_live_open" => {
                let a: OpenArgs = parse(args)?;
                let resolved = document_path(&a.path)?;
                let app = self.attach(true)?;
                // Lookup without swallowing enumeration errors: only a missing document allows Open.
                let docs = app.object("Documents")?;
                let mut existing = None;
                for index in 1..=docs.get("Count")?.integer()? {
                    let d = docs.method("Item", vec![Variant::int(index)])?.dispatch()?;
                    let full = d.get("FullName")?.string()?;
                    if Path::new(&full).is_absolute() && same_path(&resolved, Path::new(&full)) {
                        if existing.is_some() {
                            bail!("Multiple open documents match this path");
                        }
                        existing = Some(d);
                    }
                }
                let reused = existing.is_some();
                let doc = if let Some(doc) = existing {
                    doc
                } else {
                    let previous = app.get("AutomationSecurity")?.integer()?;
                    app.put("AutomationSecurity", Variant::int(3))?; // msoAutomationSecurityForceDisable
                    let opened = docs.named_method(
                        "Open",
                        &[
                            "FileName",
                            "ConfirmConversions",
                            "ReadOnly",
                            "AddToRecentFiles",
                            "Visible",
                            "NoEncodingDialog",
                        ],
                        vec![
                            Variant::text(&automation_path(&resolved)),
                            Variant::boolean(false),
                            Variant::boolean(a.read_only),
                            Variant::boolean(false),
                            Variant::boolean(a.visible),
                            Variant::boolean(true),
                        ],
                    );
                    let restored = app.put("AutomationSecurity", Variant::int(previous));
                    match (opened, restored) {
                        (Ok(value), Ok(())) => value.dispatch()?,
                        (Err(error), Ok(())) => return Err(error),
                        (Ok(_), Err(error)) => {
                            return Err(error).context(
                                "Document opened but Word AutomationSecurity could not be restored",
                            );
                        }
                        (Err(open_error), Err(restore_error)) => bail!(
                            "Opening document failed: {open_error:#}; restoring Word AutomationSecurity also failed: {restore_error:#}"
                        ),
                    }
                };
                if a.visible {
                    app.put("Visible", Variant::boolean(true))?;
                    doc.method("Activate", vec![])?;
                }
                Ok(
                    json!({"path":doc.get("FullName")?.string()?,"reused":reused,"read_only":doc.get("ReadOnly")?.bool_value()?,"visible":app.get("Visible")?.bool_value()?}),
                )
            }
            "word_live_read" => {
                let a: ReadArgs = parse(args)?;
                let (_, doc) = self.document(&a.path)?;
                let content = doc.object("Content")?;
                let start = a.start.unwrap_or(content.get("Start")?.integer()?);
                let end = a.end.unwrap_or(content.get("End")?.integer()?);
                check_range(&content, start, end)?;
                let range = doc
                    .method("Range", vec![Variant::int(start), Variant::int(end)])?
                    .dispatch()?;
                Ok(
                    json!({"path":doc.get("FullName")?.string()?,"text":range.get("Text")?.string()?,"start":start,"end":end,"saved":doc.get("Saved")?.bool_value()?}),
                )
            }
            "word_live_replace_text" => {
                let a: ReplaceArgs = parse(args)?;
                let (app, doc) = self.document(&a.path)?;
                let matches = literal_matches(&doc, &a.find, a.all)?;
                if matches.is_empty() {
                    return Ok(json!({"replacements":0,"saved":doc.get("Saved")?.bool_value()?}));
                }
                mutation(
                    &app,
                    &doc,
                    a.tracked_changes,
                    "Word MCP replace text",
                    || {
                        for &(start, end) in matches.iter().rev() {
                            doc.method("Range", vec![Variant::int(start), Variant::int(end)])?
                                .dispatch()?
                                .put("Text", Variant::text(&a.replacement))?;
                        }
                        Ok(())
                    },
                )?;
                Ok(json!({"replacements":matches.len(),"saved":doc.get("Saved")?.bool_value()?}))
            }
            "word_live_insert_text" => {
                let a: InsertArgs = parse(args)?;
                let (app, doc) = self.document(&a.path)?;
                check_range(&doc.object("Content")?, a.position, a.position)?;
                if !a.text.is_empty() {
                    mutation(
                        &app,
                        &doc,
                        a.tracked_changes,
                        "Word MCP insert text",
                        || {
                            doc.method(
                                "Range",
                                vec![Variant::int(a.position), Variant::int(a.position)],
                            )?
                            .dispatch()?
                            .put("Text", Variant::text(&a.text))
                        },
                    )?;
                }
                Ok(
                    json!({"inserted_utf16":a.text.encode_utf16().count(),"saved":doc.get("Saved")?.bool_value()?}),
                )
            }
            "word_live_save" => {
                let a: PathArgs = parse(args)?;
                let (_, doc) = self.document(&a.path)?;
                writable(&doc)?;
                doc.method("Save", vec![])?;
                Ok(
                    json!({"path":doc.get("FullName")?.string()?,"saved":doc.get("Saved")?.bool_value()?}),
                )
            }
            "word_live_export_pdf" => {
                let a: ExportArgs = parse(args)?;
                let (_, doc) = self.document(&a.path)?;
                let output = output_path(&a.output_path, a.overwrite)?;
                let temporary = tempfile::Builder::new()
                    .prefix(".word-mcp-")
                    .suffix(".pdf")
                    .tempfile_in(output.parent().context("Missing PDF parent")?)?
                    .into_temp_path();
                doc.method(
                    "ExportAsFixedFormat",
                    vec![
                        Variant::text(&automation_path(&temporary)),
                        Variant::int(17),
                        Variant::boolean(false),
                    ],
                )?;
                if temporary.metadata()?.len() == 0 {
                    bail!("Word reported export success but the PDF is empty");
                }
                if a.overwrite {
                    temporary.persist(&output)
                } else {
                    temporary.persist_noclobber(&output)
                }
                .context("Could not publish the exported PDF")?;
                Ok(json!({"output_path":automation_path(&output),"bytes":output.metadata()?.len()}))
            }
            "word_live_undo" => {
                let a: UndoArgs = parse(args)?;
                let (_, doc) = self.document(&a.path)?;
                writable(&doc)?;
                let undone = doc
                    .method("Undo", vec![Variant::int(a.count)])?
                    .bool_value()?;
                Ok(
                    json!({"undone":undone,"requested_count":a.count,"saved":doc.get("Saved")?.bool_value()?}),
                )
            }
            "word_live_view" => {
                let a: PathArgs = parse(args)?;
                let (app, doc) = self.document(&a.path)?;
                app.put("Visible", Variant::boolean(true))?;
                doc.method("Activate", vec![])?;
                Ok(json!({"visible":true,"path":doc.get("FullName")?.string()?}))
            }
            _ => bail!("Unknown live Word tool: {name}"),
        }
    }
}

// Let Word return authoritative ranges: fields can shift plain-text offsets.
fn literal_matches(doc: &Dispatch, text: &str, all: bool) -> Result<Vec<(i32, i32)>> {
    // Word returns manual line breaks as VT; accept LF as the same search input.
    let literal = text.replace('\n', "\u{b}");
    let content = doc.object("Content")?;
    let mut position = content.get("Start")?.integer()?;
    let limit = content.get("End")?.integer()?;
    let mut matches = Vec::new();
    while position < limit {
        let range = doc
            .method("Range", vec![Variant::int(position), Variant::int(limit)])?
            .dispatch()?;
        let find = range.object("Find")?;
        find.method("ClearFormatting", vec![])?;
        find.put("IgnorePunct", Variant::boolean(false))?;
        find.put("IgnoreSpace", Variant::boolean(false))?;
        find.put("MatchPrefix", Variant::boolean(false))?;
        find.put("MatchSuffix", Variant::boolean(false))?;
        let found = find
            .named_method(
                "Execute",
                &[
                    "FindText",
                    "MatchCase",
                    "MatchWholeWord",
                    "MatchWildcards",
                    "MatchSoundsLike",
                    "MatchAllWordForms",
                    "Forward",
                    "Wrap",
                    "Format",
                ],
                vec![
                    Variant::text(&word_find_text(text)),
                    Variant::boolean(true),
                    Variant::boolean(false),
                    Variant::boolean(false),
                    Variant::boolean(false),
                    Variant::boolean(false),
                    Variant::boolean(true),
                    Variant::int(0),
                    Variant::boolean(false),
                ],
            )?
            .bool_value()?;
        if !found {
            break;
        }
        let start = range.get("Start")?.integer()?;
        let end = range.get("End")?.integer()?;
        if end <= position {
            bail!("Word search made no forward progress");
        }
        if range.get("Text")?.string()? == literal {
            matches.push((start, end));
            if !all {
                break;
            }
        }
        position = end;
    }
    Ok(matches)
}

fn automation_path(path: &Path) -> String {
    let s = path.to_string_lossy();
    if let Some(unc) = s.strip_prefix(r"\\?\UNC\") {
        format!(r"\\{unc}")
    } else {
        s.strip_prefix(r"\\?\").unwrap_or(&s).to_owned()
    }
}
fn same_path(a: &Path, b: &Path) -> bool {
    let normalized = b.canonicalize().unwrap_or_else(|_| b.to_owned());
    automation_path(a).to_lowercase() == automation_path(&normalized).to_lowercase()
}
fn check_range(content: &Dispatch, start: i32, end: i32) -> Result<()> {
    let first = content.get("Start")?.integer()?;
    let last = content.get("End")?.integer()?;
    if start < first || end > last || start > end {
        bail!("Range is outside the document main story ({first}..{last})");
    }
    Ok(())
}
fn writable(doc: &Dispatch) -> Result<()> {
    if doc.get("ReadOnly")?.bool_value()? {
        bail!("Document is read-only");
    }
    Ok(())
}
fn mutation(
    app: &Dispatch,
    doc: &Dispatch,
    tracking: Option<bool>,
    label: &str,
    edit: impl FnOnce() -> Result<()>,
) -> Result<()> {
    writable(doc)?;
    // Word's custom undo records belong to the active document. Explicitly activate the target.
    doc.method("Activate", vec![])?;
    let previous = doc.get("TrackRevisions")?.bool_value()?;
    let undo = app.object("UndoRecord")?;
    undo.method("StartCustomRecord", vec![Variant::text(label)])?;
    let result = (|| {
        if let Some(enabled) = tracking {
            doc.put("TrackRevisions", Variant::boolean(enabled))?;
        }
        edit()
    })();
    let restore = if tracking.is_some() {
        doc.put("TrackRevisions", Variant::boolean(previous))
    } else {
        Ok(())
    };
    let end = undo.method("EndCustomRecord", vec![]);
    // Always attempt both cleanup calls, including when an edit fails partway through.
    result.context("Live edit failed; prior partial edits may remain and can be undone")?;
    restore.context("Edit completed but Track Changes setting could not be restored")?;
    end.context("Edit completed but Word undo record could not be ended")?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    struct TestWord(Dispatch);
    impl Drop for TestWord {
        fn drop(&mut self) {
            let _ = self.0.method("Quit", vec![Variant::int(0)]);
        }
    }

    #[test]
    fn variant_roundtrip() {
        assert_eq!(Variant::int(42).integer().unwrap(), 42);
        assert!(Variant::boolean(true).bool_value().unwrap());
        assert_eq!(Variant::text("hello 🦀").string().unwrap(), "hello 🦀");
    }
    #[test]
    fn windows_path_normalization() {
        assert_eq!(
            automation_path(Path::new(r"\\?\C:\test.docx")),
            r"C:\test.docx"
        );
        assert_eq!(
            automation_path(Path::new(r"\\?\UNC\server\share\test.docx")),
            r"\\server\share\test.docx"
        );
    }

    /// Exercises only a fresh document and a separate test-created Word instance.
    #[test]
    #[ignore = "Requires installed desktop Word and creates a temporary visible Word window"]
    fn live_word_roundtrip() {
        if unsafe { CLSIDFromProgID(w!("Word.Application")) }.is_err() {
            eprintln!("Skipping: desktop Word is not installed");
            return;
        }
        std::thread::spawn(|| -> Result<()> {
            unsafe { CoInitializeEx(None,COINIT_APARTMENTTHREADED) }.ok()?;
            let _apartment=Apartment;
            let clsid=unsafe { CLSIDFromProgID(w!("Word.Application")) }?;
            let app=Dispatch::from_windows(unsafe { CoCreateInstance::<_,IDispatch>(&raw const clsid,None,CLSCTX_LOCAL_SERVER) }?);
            // This guard only owns the instance created above, never a user's attached Word.
            let _word_guard=TestWord(app.clone());
            let directory=tempfile::tempdir()?;
            let path=directory.path().join("live-test.docx");
            let path=automation_path(&path);
            let pdf=automation_path(&directory.path().join("live-test.pdf"));
            let doc=app.object("Documents")?.method("Add",vec![])?.dispatch()?;
            doc.object("Content")?.put("Text",Variant::text("alpha 🦀 beta"))?;
            doc.method("SaveAs2",vec![Variant::text(&path),Variant::int(16)])?;
            let mut state=State { app:Some(app.clone()) };
            assert_eq!(state.call("word_live_status",&json!({}))?["running"],true);
            assert_eq!(state.call("word_live_open",&json!({"path":path,"visible":false}))?["reused"],true);
            assert!(state.call("word_live_read",&json!({"path":path}))?["text"].as_str().unwrap().contains("alpha 🦀 beta"));
            state.call("word_live_insert_text",&json!({"path":path,"position":0,"text":"PREFIX ","tracked_changes":false}))?;
            assert_eq!(state.call("word_live_replace_text",&json!({"path":path,"find":"beta","replacement":"delta","all":true,"tracked_changes":false}))?["replacements"],1);
            let edited=state.call("word_live_read",&json!({"path":path}))?;
            assert!(edited["text"].as_str().unwrap().contains("PREFIX alpha 🦀 delta"));
            assert_eq!(edited["saved"],false);
            assert_eq!(state.call("word_live_undo",&json!({"path":path}))?["undone"],true);
            assert!(state.call("word_live_read",&json!({"path":path}))?["text"].as_str().unwrap().contains("PREFIX alpha 🦀 beta"));
            let field_range=doc.method("Range",vec![Variant::int(0),Variant::int(0)])?.dispatch()?;
            let mut field_argument=Variant::empty();
            unsafe { (*field_argument.0.Anonymous.Anonymous).vt=VT_DISPATCH; (*field_argument.0.Anonymous.Anonymous).Anonymous.pdispVal=ManuallyDrop::new(Some(IDispatch::from_raw(field_range.0.clone().into_raw()))); }
            doc.object("Fields")?.method("Add",vec![field_argument,Variant::int(-1),Variant::text("DATE"),Variant::boolean(true)])?;
            assert_eq!(state.call("word_live_replace_text",&json!({"path":path,"find":"beta","replacement":"gamma"}))?["replacements"],1);
            assert!(state.call("word_live_read",&json!({"path":path}))?["text"].as_str().unwrap().contains("alpha 🦀 gamma"));
            state.call("word_live_insert_text",&json!({"path":path,"position":0,"text":"^p "}))?;
            assert_eq!(state.call("word_live_replace_text",&json!({"path":path,"find":"^p","replacement":"[literal]"}))?["replacements"],1);
            assert!(state.call("word_live_read",&json!({"path":path}))?["text"].as_str().unwrap().starts_with("[literal] "));
            state.call("word_live_insert_text",&json!({"path":path,"position":0,"text":"left\u{b}right "}))?;
            assert_eq!(state.call("word_live_replace_text",&json!({"path":path,"find":"left\nright","replacement":"LF"}))?["replacements"],1);
            state.call("word_live_insert_text",&json!({"path":path,"position":0,"text":"left\u{b}right "}))?;
            assert_eq!(state.call("word_live_replace_text",&json!({"path":path,"find":"left\u{b}right","replacement":"VT"}))?["replacements"],1);
            let original=doc.get("TrackRevisions")?.bool_value()?;
            state.call("word_live_insert_text",&json!({"path":path,"position":0,"text":"TRACKED ","tracked_changes":true}))?;
            assert_eq!(doc.get("TrackRevisions")?.bool_value()?,original);
            assert_eq!(state.call("word_live_save",&json!({"path":path}))?["saved"],true);
            state.call("word_live_export_pdf",&json!({"path":path,"output_path":pdf}))?;
            assert!(std::fs::read(&pdf)?.starts_with(b"%PDF-"));
            assert!(state.call("word_live_export_pdf",&json!({"path":path,"output_path":pdf})).is_err());
            state.call("word_live_export_pdf",&json!({"path":path,"output_path":pdf,"overwrite":true}))?;
            assert_eq!(state.call("word_live_view",&json!({"path":path}))?["visible"],true);
            doc.method("Close",vec![Variant::int(0)])?;
            // Now test opening from disk (including macro-security restoration), then close ours.
            let original_security=app.get("AutomationSecurity")?.integer()?;
            assert_eq!(state.call("word_live_open",&json!({"path":path,"visible":false}))?["reused"],false);
            assert_eq!(app.get("AutomationSecurity")?.integer()?,original_security);
            state.document(&path)?.1.method("Close",vec![Variant::int(0)])?;
            Ok(())
        }).join().expect("Word test thread panicked").expect("Word automation roundtrip failed");
    }
}

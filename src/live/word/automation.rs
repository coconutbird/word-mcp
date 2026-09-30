//! Late-bound Word automation through `IDispatch`.
//!
//! `IDispatch` is declared with `cppvtable-com`, and every Word object is owned by a
//! [`ComPtr`], so reference counting and `QueryInterface` go through cppvtable. The
//! `windows` crate supplies only the OLE functions and the `VARIANT` ABI types.
//! Interface pointers that it hands out cross into cppvtable ownership in [`adopt`]
//! and [`Variant::into_object`], without an extra reference.
use std::{mem::ManuallyDrop, time::Duration};

use anyhow::{Context, Result, bail};
use cppvtable_com::{ComPtr, GUID, HRESULT, IUnknown, interface};
use windows::{
    Win32::{
        Foundation::VARIANT_BOOL,
        System::{
            Com::{
                CLSCTX_LOCAL_SERVER, CLSIDFromProgID, COINIT_APARTMENTTHREADED, CoCreateInstance,
                CoInitializeEx, CoUninitialize, DISPATCH_FLAGS, DISPATCH_METHOD,
                DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
            },
            Ole::GetActiveObject,
            Variant::{
                VARENUM, VARIANT, VARIANT_0_0_0, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_EMPTY, VT_I2,
                VT_I4, VT_INT, VT_NULL, VT_R8, VT_UNKNOWN, VariantClear,
            },
        },
        UI::WindowsAndMessaging::{
            DispatchMessageW, MSG, PM_REMOVE, PeekMessageW, TranslateMessage,
        },
    },
    core::{BSTR, Interface as _, PCWSTR, w},
};

/// `LOCALE_USER_DEFAULT` would localize member names; Word's names are English.
const LOCALE_EN_US: u32 = 0x409;
/// `DISPID_PROPERTYPUT`: the named argument that carries a property's new value.
const DISPID_PROPERTYPUT: i32 = -3;
/// `MK_E_UNAVAILABLE`: no running object is registered for the class.
const MK_E_UNAVAILABLE: u32 = 0x8004_01E3;
/// `RPC_E_CALL_REJECTED` and `RPC_E_SERVERCALL_RETRYLATER`: Word is busy and did not
/// execute the call, so it is safe to retry.
const BUSY: [u32; 2] = [0x8001_0001, 0x8001_010A];
/// How often, and how long apart, a busy rejection is retried.
const BUSY_RETRIES: u32 = 50;
const BUSY_DELAY: Duration = Duration::from_millis(100);

/// The OLE Automation interface through which Word exposes its object model.
#[interface(abi = com, iid = "00020400-0000-0000-C000-000000000046")]
pub(in crate::live) unsafe trait IDispatch {
    /// Report whether the object provides type information.
    ///
    /// # Safety
    /// `count` must be aligned and writable for one `u32`.
    unsafe fn GetTypeInfoCount(&self, count: *mut u32) -> HRESULT;

    /// Return the object's `ITypeInfo`.
    ///
    /// # Safety
    /// `info` must be aligned and writable for one interface pointer.
    unsafe fn GetTypeInfo(
        &self,
        index: u32,
        locale: u32,
        info: *mut Option<ComPtr<IUnknown>>,
    ) -> HRESULT;

    /// Map member and named-argument names to dispatch identifiers.
    ///
    /// # Safety
    /// `iid` must point to `IID_NULL`; `names` must point to `count` NUL-terminated
    /// UTF-16 strings, and `ids` must be writable for `count` identifiers.
    unsafe fn GetIDsOfNames(
        &self,
        iid: *const GUID,
        names: *const PCWSTR,
        count: u32,
        locale: u32,
        ids: *mut i32,
    ) -> HRESULT;

    /// Invoke a member.
    ///
    /// # Safety
    /// `iid` must point to `IID_NULL`; `params` must describe valid argument arrays in
    /// reverse order; `result`, `exception`, and `argument` must be writable, and
    /// `result` must hold an initialized `VARIANT` that the call may overwrite.
    unsafe fn Invoke(
        &self,
        id: i32,
        iid: *const GUID,
        locale: u32,
        flags: u16,
        params: *const DISPPARAMS,
        result: *mut VARIANT,
        exception: *mut EXCEPINFO,
        argument: *mut u32,
    ) -> HRESULT;
}

/// An owned reference to a Word automation object, confined to the STA thread.
pub(in crate::live) type Object = ComPtr<IDispatch>;

impl IDispatch {
    fn ids(&self, names: &[&str]) -> Result<Vec<i32>> {
        let wide: Vec<Vec<u16>> = names
            .iter()
            .map(|name| name.encode_utf16().chain([0]).collect())
            .collect();
        let pointers: Vec<PCWSTR> = wide.iter().map(|name| PCWSTR(name.as_ptr())).collect();
        let mut ids = vec![0; names.len()];
        // SAFETY: `pointers` holds `names.len()` NUL-terminated strings kept alive by
        // `wide`, and `ids` has room for one identifier per name.
        unsafe {
            self.GetIDsOfNames(
                &GUID::zeroed(),
                pointers.as_ptr(),
                u32::try_from(pointers.len())?,
                LOCALE_EN_US,
                ids.as_mut_ptr(),
            )
        }
        .ok()
        .with_context(|| format!("Word has no member {}", names.join("/")))?;
        Ok(ids)
    }

    fn invoke(
        &self,
        name: &str,
        flags: DISPATCH_FLAGS,
        id: i32,
        mut arguments: Vec<Variant>,
        mut named: Vec<i32>,
    ) -> Result<Variant> {
        // IDispatch takes arguments last-to-first.
        arguments.reverse();
        named.reverse();
        let params = DISPPARAMS {
            rgvarg: arguments.as_mut_ptr().cast(),
            rgdispidNamedArgs: named.as_mut_ptr(),
            cArgs: u32::try_from(arguments.len())?,
            cNamedArgs: u32::try_from(named.len())?,
        };
        for _ in 0..BUSY_RETRIES {
            let mut result = Variant::empty();
            let mut exception = EXCEPINFO::default();
            let mut argument = 0;
            // SAFETY: `params` points into `arguments`/`named`, which outlive the call
            // (`Variant` is a transparent `VARIANT`); the out-parameters are locals.
            let status = unsafe {
                self.Invoke(
                    id,
                    &GUID::zeroed(),
                    LOCALE_EN_US,
                    flags.0,
                    &raw const params,
                    &raw mut result.0,
                    &raw mut exception,
                    &raw mut argument,
                )
            };
            let detail = take_exception(&mut exception);
            if status.is_ok() {
                return Ok(result);
            }
            if !BUSY.contains(&status.0.cast_unsigned()) {
                let code = if detail.code == 0 {
                    status.0
                } else {
                    detail.code
                };
                let message = detail.description.unwrap_or_else(|| status.message());
                bail!(
                    "Word {name} failed: {message} (HRESULT {:#010X}, argument {argument})",
                    code.cast_unsigned()
                );
            }
            pump();
            std::thread::sleep(BUSY_DELAY);
        }
        bail!("Word stayed busy during {name}; dismiss any open Word dialog and retry")
    }

    /// Read a property.
    pub(in crate::live) fn get(&self, name: &str) -> Result<Variant> {
        let id = self.ids(&[name])?[0];
        self.invoke(name, DISPATCH_PROPERTYGET, id, vec![], vec![])
    }

    /// Assign a property.
    pub(in crate::live) fn put(&self, name: &str, value: Variant) -> Result<()> {
        let id = self.ids(&[name])?[0];
        self.invoke(
            name,
            DISPATCH_PROPERTYPUT,
            id,
            vec![value],
            vec![DISPID_PROPERTYPUT],
        )
        .map(drop)
    }

    /// Call a method with positional arguments.
    pub(in crate::live) fn call(&self, name: &str, arguments: Vec<Variant>) -> Result<Variant> {
        let id = self.ids(&[name])?[0];
        self.invoke(name, DISPATCH_METHOD, id, arguments, vec![])
    }

    /// Call a method with named arguments, paired with `arguments` in order.
    pub(in crate::live) fn call_named(
        &self,
        name: &str,
        names: &[&str],
        arguments: Vec<Variant>,
    ) -> Result<Variant> {
        debug_assert_eq!(names.len(), arguments.len());
        let mut all = vec![name];
        all.extend_from_slice(names);
        let mut ids = self.ids(&all)?;
        let id = ids.remove(0);
        self.invoke(name, DISPATCH_METHOD, id, arguments, ids)
    }

    /// Read an object-valued property.
    pub(in crate::live) fn object(&self, name: &str) -> Result<Object> {
        self.get(name)?
            .into_object()?
            .with_context(|| format!("Word returned no {name} object"))
    }

    /// Read an integer property.
    pub(in crate::live) fn int(&self, name: &str) -> Result<i32> {
        self.get(name)?
            .int()
            .with_context(|| format!("Word {name}"))
    }

    /// Read a boolean property.
    pub(in crate::live) fn flag(&self, name: &str) -> Result<bool> {
        self.get(name)?
            .flag()
            .with_context(|| format!("Word {name}"))
    }

    /// Read a text property.
    pub(in crate::live) fn string(&self, name: &str) -> Result<String> {
        self.get(name)?
            .string()
            .with_context(|| format!("Word {name}"))
    }
}

struct ExceptionDetail {
    code: i32,
    description: Option<String>,
}

/// Fill in and release an `EXCEPINFO`'s strings, returning its useful parts.
fn take_exception(exception: &mut EXCEPINFO) -> ExceptionDetail {
    // SAFETY: The server filled `exception` (or left it zeroed); the deferred fill-in
    // callback is its own, and each BSTR is released exactly once here.
    unsafe {
        if let Some(fill) = exception.pfnDeferredFillIn {
            let _ = fill(exception);
        }
        let description = exception.bstrDescription.to_string();
        ManuallyDrop::drop(&mut exception.bstrSource);
        ManuallyDrop::drop(&mut exception.bstrDescription);
        ManuallyDrop::drop(&mut exception.bstrHelpFile);
        ExceptionDetail {
            code: exception.scode,
            description: (!description.trim().is_empty()).then(|| description.trim().to_owned()),
        }
    }
}

/// Transfer an interface pointer from the `windows` crate into cppvtable ownership
/// and query it for `IDispatch`.
fn adopt(unknown: windows::core::IUnknown) -> Result<Object> {
    // SAFETY: `into_raw` yields the one reference that `unknown` owned.
    let unknown = unsafe { ComPtr::<IUnknown>::from_raw_unchecked(unknown.into_raw()) };
    unknown
        .cast::<IDispatch>()
        .context("Word object does not support automation")
}

/// An owned `VARIANT`. `windows` does not clear `VARIANT`s, so this does.
#[repr(transparent)]
pub(in crate::live) struct Variant(VARIANT);

impl Variant {
    fn empty() -> Self {
        Self(VARIANT::default())
    }

    /// Build a variant of type `vt`, then let `fill` set the matching union member.
    fn with(vt: VARENUM, fill: impl FnOnce(&mut VARIANT_0_0_0)) -> Self {
        let mut variant = Self::empty();
        // SAFETY: The tag and the member `fill` writes are set together, so the
        // variant never claims a member that is not initialized.
        unsafe {
            let inner = &mut *variant.0.Anonymous.Anonymous;
            fill(&mut inner.Anonymous);
            inner.vt = vt;
        }
        variant
    }

    /// Pass a Word object as an argument; the variant owns one added reference.
    pub(in crate::live) fn object(object: &Object) -> Self {
        let raw = object.clone().into_raw();
        Self::with(VT_DISPATCH, |member| {
            // SAFETY: `raw` carries the reference added by `clone`; the `windows`
            // wrapper takes it over and `VariantClear` releases it.
            let dispatch = unsafe { windows::Win32::System::Com::IDispatch::from_raw(raw) };
            member.pdispVal = ManuallyDrop::new(Some(dispatch));
        })
    }

    fn vt(&self) -> VARENUM {
        // SAFETY: The tag is always initialized.
        unsafe { self.0.Anonymous.Anonymous.vt }
    }

    pub(in crate::live) fn int(&self) -> Result<i32> {
        // SAFETY: Each arm reads the member that the tag selects.
        unsafe {
            let inner = &self.0.Anonymous.Anonymous.Anonymous;
            match self.vt() {
                VT_I4 | VT_INT => Ok(inner.lVal),
                VT_I2 => Ok(i32::from(inner.iVal)),
                other => bail!("expected an integer, got VARIANT type {}", other.0),
            }
        }
    }

    /// Word reports sizes as `Single`, so accept both floating-point types.
    #[cfg(test)]
    pub(in crate::live) fn number(&self) -> Result<f64> {
        use windows::Win32::System::Variant::VT_R4;
        // SAFETY: Each arm reads the member that the tag selects.
        unsafe {
            let inner = &self.0.Anonymous.Anonymous.Anonymous;
            match self.vt() {
                VT_R8 => Ok(inner.dblVal),
                VT_R4 => Ok(f64::from(inner.fltVal)),
                _ => self.int().map(f64::from),
            }
        }
    }

    pub(in crate::live) fn flag(&self) -> Result<bool> {
        // SAFETY: Each arm reads the member that the tag selects.
        unsafe {
            match self.vt() {
                VT_BOOL => Ok(self.0.Anonymous.Anonymous.Anonymous.boolVal.0 != 0),
                other => bail!("expected a boolean, got VARIANT type {}", other.0),
            }
        }
    }

    pub(in crate::live) fn string(&self) -> Result<String> {
        // SAFETY: Each arm reads the member that the tag selects.
        unsafe {
            match self.vt() {
                VT_BSTR => Ok(self.0.Anonymous.Anonymous.Anonymous.bstrVal.to_string()),
                other => bail!("expected text, got VARIANT type {}", other.0),
            }
        }
    }

    /// Take the object out of the variant; `None` for `Nothing`/empty/null.
    pub(in crate::live) fn into_object(mut self) -> Result<Option<Object>> {
        let vt = self.vt();
        // SAFETY: Each arm moves out the interface member that the tag selects and
        // then marks the variant empty, so `Drop` does not release it again.
        let owned = unsafe {
            let inner = &mut *self.0.Anonymous.Anonymous;
            match vt {
                VT_EMPTY | VT_NULL => return Ok(None),
                VT_DISPATCH => {
                    let dispatch = ManuallyDrop::take(&mut inner.Anonymous.pdispVal);
                    inner.vt = VT_EMPTY;
                    dispatch.map(|dispatch| {
                        // The `windows` wrapper's single reference moves into the ComPtr.
                        ComPtr::<IDispatch>::from_raw_unchecked(dispatch.into_raw())
                    })
                }
                VT_UNKNOWN => {
                    let unknown = ManuallyDrop::take(&mut inner.Anonymous.punkVal);
                    inner.vt = VT_EMPTY;
                    return unknown.map(adopt).transpose();
                }
                other => bail!("expected an object, got VARIANT type {}", other.0),
            }
        };
        Ok(owned)
    }
}

impl Drop for Variant {
    fn drop(&mut self) {
        // SAFETY: The variant is initialized and owns whatever its tag selects.
        let _ = unsafe { VariantClear(&raw mut self.0) };
    }
}

impl From<i32> for Variant {
    fn from(value: i32) -> Self {
        Self::with(VT_I4, |member| member.lVal = value)
    }
}

impl From<f64> for Variant {
    fn from(value: f64) -> Self {
        Self::with(VT_R8, |member| member.dblVal = value)
    }
}

impl From<bool> for Variant {
    fn from(value: bool) -> Self {
        Self::with(VT_BOOL, |member| {
            member.boolVal = VARIANT_BOOL(if value { -1 } else { 0 });
        })
    }
}

impl From<&str> for Variant {
    fn from(value: &str) -> Self {
        // `VariantClear` frees the BSTR.
        Self::with(VT_BSTR, |member| {
            member.bstrVal = ManuallyDrop::new(BSTR::from(value));
        })
    }
}

/// Dispatch pending window messages; an STA must pump for Word's callbacks.
pub(in crate::live) fn pump() {
    let mut message = MSG::default();
    // SAFETY: `message` is a local, and messages are dispatched on the thread that
    // retrieved them.
    unsafe {
        while PeekMessageW(&raw mut message, None, 0, 0, PM_REMOVE).as_bool() {
            let _ = TranslateMessage(&raw const message);
            DispatchMessageW(&raw const message);
        }
    }
}

/// A single-threaded COM apartment on the current thread, left on drop. Drop every
/// [`Object`] first.
pub(in crate::live) struct Apartment(());

impl Apartment {
    pub(in crate::live) fn enter() -> Result<Self> {
        // SAFETY: Called once per thread; paired with `CoUninitialize` in `Drop`.
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }
            .ok()
            .context("cannot initialize the Word COM apartment")?;
        Ok(Self(()))
    }
}

impl Drop for Apartment {
    fn drop(&mut self) {
        // SAFETY: Balances the successful `CoInitializeEx` in `enter`.
        unsafe { CoUninitialize() };
    }
}

/// Word's registered class, or `None` when desktop Word is not installed.
pub(in crate::live) fn word_class() -> Option<GUID> {
    // SAFETY: The ProgID is a static NUL-terminated string.
    unsafe { CLSIDFromProgID(w!("Word.Application")) }.ok()
}

/// The running Word instance, or `None` when Word is not running.
pub(in crate::live) fn running_word(class: &GUID) -> Result<Option<Object>> {
    let mut unknown = None;
    // SAFETY: `class` and `unknown` are valid for the call.
    match unsafe { GetActiveObject(class, None, &raw mut unknown) } {
        Ok(()) => adopt(unknown.context("running Word returned no object")?).map(Some),
        Err(error) if error.code().0.cast_unsigned() == MK_E_UNAVAILABLE => Ok(None),
        Err(error) => Err(error).context("cannot attach to the running Microsoft Word"),
    }
}

/// Start a new Word instance.
pub(in crate::live) fn launch_word(class: &GUID) -> Result<Object> {
    // SAFETY: `class` is Word's registered class and COM is initialized.
    let unknown =
        unsafe { CoCreateInstance::<_, windows::core::IUnknown>(class, None, CLSCTX_LOCAL_SERVER) }
            .context("cannot launch Microsoft Word")?;
    adopt(unknown)
}

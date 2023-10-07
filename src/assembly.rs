use std::{
    ffi::{c_long, c_void},
    io,
    ptr::null_mut,
};

use windows::{
    core::{IUnknown_Vtbl, BSTR, HRESULT},
    Win32::System::Com::SAFEARRAY,
    Win32::System::Variant::VARIANT,
};

use crate::methodinfo::MethodInfo;

#[repr(C)]
pub struct Assembly {
    pub vtable: *const AssemblyVtbl,
}

impl Assembly {
    pub fn get_entrypoint(&self) -> Result<*mut MethodInfo, String> {
        let mut method_info: *mut MethodInfo = null_mut();
        let res = unsafe {
            ((*self.vtable).get_EntryPoint)(self as *const _ as *mut _, &mut method_info)
        };

        match res.0 {
            0 => Ok(method_info),
            _ => Err(format!(
                "Couldn't find the entrypoint, {:?}",
                io::Error::last_os_error()
            )),
        }
    }

    //pub fn get_type(&self, name: &str) -> Result<*mut Type, String> {
    //    let dw = BSTR::from(name);
    //
    //    let mut type_ptr: *mut Type = null_mut();
    //    let hr = unsafe {
    //        ((*self.vtable).GetType_2)(
    //            self as *const _ as *mut _,
    //            dw.into_raw() as *mut _,
    //            &mut type_ptr,
    //        )
    //    };
    //
    //    if hr.is_err() {
    //        return Err(format!(
    //            "Error while retrieving type `{}`: 0x{:x}",
    //            name, hr.0
    //        ));
    //    }
    //
    //    if type_ptr.is_null() {
    //        return Err(format!("Could not retrieve type `{}`", name));
    //    }
    //
    //    Ok(type_ptr)
    //}

    pub fn create_instance(&self, name: &str) -> Result<VARIANT, String> {
        let dw = BSTR::from(name);

        let mut instance: VARIANT = VARIANT::default();
        let hr = unsafe {
            ((*self.vtable).CreateInstance)(
                self as *const _ as *mut _,
                dw.into_raw() as *mut _,
                &mut instance,
            )
        };

        if hr.is_err() {
            return Err(format!(
                "Error while creating instance of `{}`: 0x{:x}",
                name, hr.0
            ));
        }

        Ok(instance)
    }

    //pub fn get_types(&self) -> Result<Vec<*mut Type>, String> {
    //    let mut results: Vec<*mut Type> = vec![];
    //
    //    let mut safe_array_ptr: *mut SAFEARRAY = unsafe { SafeArrayCreateVector(VT_UNKNOWN, 0, 0) };
    //
    //    let hr = unsafe {
    //        ((*((*self).vtable)).GetTypes)(&self as *const _ as *mut _, &mut safe_array_ptr)
    //    };
    //
    //    if hr.is_err() {
    //        return Err(format!("Error while retrieving types: 0x{:x}", hr.0));
    //    }
    //
    //    let ubound = unsafe { SafeArrayGetUBound(safe_array_ptr, 1) }.unwrap_or(0);
    //
    //    for i in 0..ubound {
    //        let indices: [i32; 1] = [i as _];
    //        let mut variant: *mut Type = null_mut();
    //        let pv = &mut variant as *mut _ as *mut c_void;
    //
    //        match unsafe { SafeArrayGetElement(safe_array_ptr, indices.as_ptr(), pv) } {
    //            Ok(_) => {}
    //            Err(e) => return Err(format!("Could not access safe array: {:?}", e.code())),
    //        }
    //
    //        unsafe { dbg!((*variant).to_string().unwrap()) };
    //
    //        if !pv.is_null() {
    //            results.push(variant)
    //        }
    //    }
    //
    //    Ok(results)
    //}
}

#[repr(C)]
#[allow(non_snake_case)]
pub struct AssemblyVtbl {
    pub parent: IUnknown_Vtbl,
    pub GetTypeInfoCount: *const c_void,
    pub GetTypeInfo: *const c_void,
    pub GetIDsOfNames: *const c_void,
    pub Invoke: *const c_void,
    pub ToString: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut u16) -> HRESULT,
    pub Equals: *const c_void,
    pub GetHashCode: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut c_long) -> HRESULT,
    pub GetType: *const c_void,
    pub get_CodeBase:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut u16) -> HRESULT,
    pub get_EscapedCodeBase:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut u16) -> HRESULT,
    pub GetName: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut c_void) -> HRESULT,
    pub GetName_2: *const c_void,
    pub get_FullName:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut u16) -> HRESULT,
    pub get_EntryPoint:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut MethodInfo) -> HRESULT,
    pub GetType_2: unsafe extern "system" fn(
        this: *mut c_void,
        name: *mut u16,
        pRetVal: *mut *mut c_void,
    ) -> HRESULT,
    pub GetType_3: *const c_void,
    pub GetExportedTypes: *const c_void,
    pub GetTypes:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut SAFEARRAY) -> HRESULT,
    pub GetManifestResourceStream: *const c_void,
    pub GetManifestResourceStream_2: *const c_void,
    pub GetFile: *const c_void,
    pub GetFiles: *const c_void,
    pub GetFiles_2: *const c_void,
    pub GetManifestResourceNames: *const c_void,
    pub GetManifestResourceInfo: *const c_void,
    pub get_Location:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut u16) -> HRESULT,
    pub get_Evidence: *const c_void,
    pub GetCustomAttributes: *const c_void,
    pub GetCustomAttributes_2: *const c_void,
    pub IsDefined: *const c_void,
    pub GetObjectData: *const c_void,
    pub add_ModuleResolve: *const c_void,
    pub remove_ModuleResolve: *const c_void,
    pub GetType_4: *const c_void,
    pub GetSatelliteAssembly: *const c_void,
    pub GetSatelliteAssembly_2: *const c_void,
    pub LoadModule: *const c_void,
    pub LoadModule_2: *const c_void,
    pub CreateInstance: unsafe extern "system" fn(
        this: *mut c_void,
        typeName: *mut u16,
        pRetVal: *mut VARIANT,
    ) -> HRESULT,
    pub CreateInstance_2: *const c_void,
    pub CreateInstance_3: *const c_void,
    pub GetLoadedModules: *const c_void,
    pub GetLoadedModules_2: *const c_void,
    pub GetModules: *const c_void,
    pub GetModules_2: *const c_void,
    pub GetModule: *const c_void,
    pub GetReferencedAssemblies: *const c_void,
    pub get_GlobalAssemblyCache: *const c_void,
}

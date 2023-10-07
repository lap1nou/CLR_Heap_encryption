use std::{
    ffi::{c_long, c_void},
    io,
    ptr::null_mut,
};

use windows::{
    core::{IUnknown_Vtbl, BSTR, GUID, HRESULT},
    Win32::System::Com::SAFEARRAY,
};

use super::assembly::Assembly;

#[repr(C)]
pub struct AppDomain {
    pub vtable: *const AppDomainVtbl,
}

impl AppDomain {
    pub fn load_assembly(&self, safe_array: *mut SAFEARRAY) -> Result<*mut Assembly, String> {
        let mut assembly: *mut Assembly = null_mut();
        let res = unsafe {
            ((*self.vtable).Load_3)(self as *const _ as *mut _, safe_array, &mut assembly)
        };

        match res.0 {
            0 => Ok(assembly),
            _ => Err(format!(
                "Couldn't load the assembly, {:?}",
                io::Error::last_os_error()
            )),
        }
    }

    pub fn load_library(&self, library: &str) -> Result<*mut Assembly, String> {
        let library_buffer = BSTR::from(library);

        let mut library_ptr: *mut Assembly = null_mut();

        let hr = unsafe {
            ((*self.vtable).Load_2)(
                self as *const _ as *mut _,
                library_buffer.into_raw() as *mut _,
                &mut library_ptr,
            )
        };

        if hr.is_err() {
            return Err(format!("Could not retrieve `{}`: {:?}", library, hr));
        }

        if library_ptr.is_null() {
            return Err(format!("Could not retrieve `{}`", library));
        }

        Ok(library_ptr)
    }

    pub const IID: GUID = GUID::from_u128(0x05F696DC_2B29_3663_AD8B_C4389CF2A713);
}

#[repr(C)]
#[allow(non_snake_case)]
pub struct AppDomainVtbl {
    pub parent: IUnknown_Vtbl,
    pub GetTypeInfoCount: *const c_void,
    pub GetTypeInfo: *const c_void,
    pub GetIDsOfNames: *const c_void,
    pub Invoke: *const c_void,
    pub ToString: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut u16) -> HRESULT,
    pub Equals: *const c_void,
    pub GetHashCode: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut c_long) -> HRESULT,
    pub GetType: *const c_void,
    pub InitializeLifetimeService: *const c_void,
    pub GetLifetimeService: *const c_void,
    pub get_Evidence: *const c_void,
    pub set_Evidence: *const c_void,
    pub get_DomainUnload: *const c_void,
    pub set_DomainUnload: *const c_void,
    pub get_AssemblyLoad: *const c_void,
    pub set_AssemblyLoad: *const c_void,
    pub get_ProcessExit: *const c_void,
    pub set_ProcessExit: *const c_void,
    pub get_TypeResolve: *const c_void,
    pub set_TypeResolve: *const c_void,
    pub get_ResourceResolve: *const c_void,
    pub set_ResourceResolve: *const c_void,
    pub get_AssemblyResolve: *const c_void,
    pub get_UnhandledException: *const c_void,
    pub set_UnhandledException: *const c_void,
    pub DefineDynamicAssembly: *const c_void,
    pub DefineDynamicAssembly_2: *const c_void,
    pub DefineDynamicAssembly_3: *const c_void,
    pub DefineDynamicAssembly_4: *const c_void,
    pub DefineDynamicAssembly_5: *const c_void,
    pub DefineDynamicAssembly_6: *const c_void,
    pub DefineDynamicAssembly_7: *const c_void,
    pub DefineDynamicAssembly_8: *const c_void,
    pub DefineDynamicAssembly_9: *const c_void,
    pub CreateInstance: *const c_void,
    pub CreateInstanceFrom: *const c_void,
    pub CreateInstance_2: *const c_void,
    pub CreateInstanceFrom_2: *const c_void,
    pub CreateInstance_3: *const c_void,
    pub CreateInstanceFrom_3: *const c_void,
    pub Load: *const c_void,
    pub Load_2: unsafe extern "system" fn(
        this: *mut c_void,
        assemblyString: *mut u16,
        pRetVal: *mut *mut Assembly,
    ) -> HRESULT,
    pub Load_3: unsafe extern "system" fn(
        this: *mut c_void,
        rawAssembly: *mut SAFEARRAY,
        pRetVal: *mut *mut Assembly,
    ) -> HRESULT,
    pub Load_4: *const c_void,
    pub Load_5: *const c_void,
    pub Load_6: *const c_void,
    pub Load_7: *const c_void,
    pub ExecuteAssembly: *const c_void,
    pub ExecuteAssembly_2: *const c_void,
    pub ExecuteAssembly_3: *const c_void,
    pub get_FriendlyName: *const c_void,
    pub get_BaseDirectory: *const c_void,
    pub get_RelativeSearchPath: *const c_void,
    pub get_ShadowCopyFiles: *const c_void,
    pub GetAssemblies: *const c_void,
    pub AppendPrivatePath: *const c_void,
    pub ClearPrivatePath: *const c_void,
    pub ClearShadowCopyPath: *const c_void,
    pub SetData: *const c_void,
    pub GetData: *const c_void,
    pub SetAppDomainPolicy: *const c_void,
    pub SetThreadPrincipal: *const c_void,
    pub SetPrincipalPolicy: *const c_void,
    pub DoCallBack: *const c_void,
    pub get_DynamicDirectory: *const c_void,
}

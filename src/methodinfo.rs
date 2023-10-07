use std::{
    ffi::{c_long, c_void},
    io,
};

use windows::{
    core::{IUnknown_Vtbl, BSTR, HRESULT},
    Win32::System::{
        Com::SAFEARRAY,
        Ole::{SafeArrayCreateVector, SafeArrayGetLBound, SafeArrayGetUBound},
        Variant::{VARIANT, VT_UNKNOWN, VT_VARIANT},
    },
};

#[derive(Debug, Copy, Clone)]
#[repr(C)]
pub struct MethodInfo {
    pub vtable: *const MethodInfoVtbl,
}

impl MethodInfo {
    pub fn invoke_assembly(&self, safe_array_final: *mut SAFEARRAY) -> Result<(), String> {
        let object: VARIANT = unsafe { std::mem::zeroed() };
        let mut return_value: VARIANT = unsafe { std::mem::zeroed() };

        let res = unsafe {
            ((*self.vtable).Invoke_3)(
                self as *const _ as *mut _,
                object,
                safe_array_final,
                &mut return_value,
            )
        };

        match res.0 {
            0 => Ok(()),
            _ => Err(format!(
                "Couldn't invoke the assembly, {:?}",
                io::Error::last_os_error()
            )),
        }
    }

    pub unsafe fn invoke_without_args(&self, instance: Option<VARIANT>) -> Result<VARIANT, String> {
        let method_args = unsafe { SafeArrayCreateVector(VT_VARIANT, 0, 0) };

        self.invoke(method_args, instance)
    }

    pub unsafe fn invoke(
        &self,
        args: *mut SAFEARRAY,
        instance: Option<VARIANT>,
    ) -> Result<VARIANT, String> {
        let args_len = get_array_length(args);
        let parameter_count = (*self).get_parameter_count()?;

        if args_len != parameter_count {
            return Err(format!(
                "Arguments do not match method signature: {} given, {} expected",
                args_len, parameter_count
            ));
        }

        let mut return_value: VARIANT = unsafe { std::mem::zeroed() };

        let object: VARIANT = match instance {
            None => unsafe { std::mem::zeroed() },
            Some(i) => i,
        };

        let hr = unsafe {
            ((*self.vtable).Invoke_3)(self as *const _ as *mut _, object, args, &mut return_value)
        };

        if hr.is_err() {
            return Err(format!("Could not invoke method: {:?}", hr));
        }

        Ok(return_value)
    }

    pub fn get_parameter_count(&self) -> Result<i32, String> {
        let mut safe_array_ptr: *mut SAFEARRAY =
            unsafe { SafeArrayCreateVector(VT_UNKNOWN, 0, 255) };

        let hr = unsafe {
            ((*self.vtable).GetParameters)(self as *const _ as *mut _, &mut safe_array_ptr)
        };

        if hr.is_err() {
            return Err(format!("Could not get parameter count: {:?}", hr));
        }

        Ok(get_array_length(safe_array_ptr))
    }

    pub fn to_string(&self) -> Result<String, String> {
        let mut buffer = BSTR::new();

        let hr = unsafe { (*self).ToString(&mut buffer as *mut _ as *mut *mut u16) };

        if hr.is_err() {
            return Err(format!("Failed while running `ToString`: {:?}", hr));
        }

        Ok(buffer.to_string())
    }

    #[inline]
    pub unsafe fn ToString(&self, pRetVal: *mut *mut u16) -> HRESULT {
        ((*self.vtable).ToString)(self as *const _ as *mut _, pRetVal)
    }
}

#[repr(C)]
#[allow(non_snake_case)]
pub struct MethodInfoVtbl {
    pub parent: IUnknown_Vtbl,
    pub GetTypeInfoCount: *const c_void,
    pub GetTypeInfo: *const c_void,
    pub GetIDsOfNames: *const c_void,
    pub Invoke: *const c_void,
    pub ToString: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut u16) -> HRESULT,
    pub Equals: *const c_void,
    pub GetHashCode: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut c_long) -> HRESULT,
    pub GetType: *const c_void,
    pub get_MemberType: *const c_void,
    pub get_name: unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut u16) -> HRESULT,
    pub get_DeclaringType: *const c_void,
    pub get_ReflectedType: *const c_void,
    pub GetCustomAttributes: *const c_void,
    pub GetCustomAttributes_2: *const c_void,
    pub IsDefined: *const c_void,
    pub GetParameters:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut SAFEARRAY) -> HRESULT,
    pub GetMethodImplementationFlags: *const c_void,
    pub get_MethodHandle: *const c_void,
    pub get_Attributes: *const c_void,
    pub get_CallingConvention: *const c_void,
    pub Invoke_2: *const c_void,
    pub get_IsPublic: *const c_void,
    pub get_IsPrivate: *const c_void,
    pub get_IsFamily: *const c_void,
    pub get_IsAssembly: *const c_void,
    pub get_IsFamilyAndAssembly: *const c_void,
    pub get_IsFamilyOrAssembly: *const c_void,
    pub get_IsStatic: *const c_void,
    pub get_IsFinal: *const c_void,
    pub get_IsVirtual: *const c_void,
    pub get_IsHideBySig: *const c_void,
    pub get_IsAbstract: *const c_void,
    pub get_IsSpecialName: *const c_void,
    pub get_IsConstructor: *const c_void,
    pub Invoke_3: unsafe extern "system" fn(
        this: *mut c_void,
        obj: VARIANT,
        parameters: *mut SAFEARRAY,
        pRetVal: *mut VARIANT,
    ) -> HRESULT,
    pub get_returnType: *const c_void,
    pub get_ReturnTypeCustomAttributes: *const c_void,
    pub GetBaseDefinition:
        unsafe extern "system" fn(this: *mut c_void, pRetVal: *mut *mut MethodInfo) -> HRESULT,
}

pub fn get_array_length(array_ptr: *mut SAFEARRAY) -> i32 {
    let upper = unsafe { SafeArrayGetUBound(array_ptr, 1) }.unwrap_or(0);
    let lower = unsafe { SafeArrayGetLBound(array_ptr, 1) }.unwrap_or(0);

    match upper - lower {
        0 => 0,
        delta => delta + 1,
    }
}

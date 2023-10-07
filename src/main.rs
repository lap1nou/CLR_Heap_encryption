use crate::appdomain::AppDomain;
use lazy_static::lazy_static;
use std::ffi::{c_int, c_void};
use std::mem::{self, ManuallyDrop};
use std::ptr::{addr_of, addr_of_mut, null_mut};
use std::sync::Mutex;
use std::{env, fs, io, ptr};
use windows::core::implement;
use windows::Win32::Foundation::{E_NOINTERFACE, E_OUTOFMEMORY, HANDLE, S_OK};
use windows::Win32::System::ClrHosting::{
    CLRCreateInstance, CLRRuntimeHost, CLSID_CLRMetaHost, CorRuntimeHost, EMemoryCriticalLevel,
    ICLRMemoryNotificationCallback, ICLRMetaHost, ICLRRuntimeHost, ICLRRuntimeInfo,
    ICorRuntimeHost, IHostControl, IHostControl_Impl, IHostMalloc, IHostMalloc_Impl,
    IHostMemoryManager, IHostMemoryManager_Impl, MALLOC_EXECUTABLE, MALLOC_THREADSAFE, MALLOC_TYPE,
};
use windows::Win32::System::Com::SAFEARRAY;
use windows::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryA};
use windows::Win32::System::Memory::{
    HeapAlloc, HeapCreate, HeapFree, VirtualAlloc, VirtualFree, VirtualProtect, VirtualQuery,
    HEAP_CREATE_ENABLE_EXECUTE, HEAP_FLAGS, HEAP_NO_SERIALIZE, MEMORY_BASIC_INFORMATION,
    MEM_RESERVE, PAGE_PROTECTION_FLAGS, VIRTUAL_ALLOCATION_TYPE, VIRTUAL_FREE_TYPE,
};
use windows::Win32::System::Ole::{SafeArrayCreateVector, SafeArrayDestroy, SafeArrayPutElement};
use windows::Win32::System::Variant::{
    VARENUM, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0, VT_ARRAY, VT_BSTR, VT_UI1, VT_VARIANT,
};
use windows_core::{w, ComInterface, IUnknown, Interface, BSTR, PCSTR};
use zeroize::Zeroize;

mod appdomain;
mod assembly;
mod methodinfo;

#[repr(C)]
pub struct UString {
    pub length: u32,
    pub maximum_length: u32,
    pub buffer: *mut c_void,
}

type FnSystemFunction032 = unsafe extern "system" fn(*const UString, *const UString) -> c_int;

#[implement(IHostControl)]
pub struct MyHostControl;

impl IHostControl_Impl for MyHostControl {
    fn GetHostManager(
        &self,
        _riid: *const ::windows_core::GUID,
        _ppobject: *mut *mut ::core::ffi::c_void,
    ) -> ::windows_core::Result<()> {
        if unsafe { _riid.as_ref().unwrap() } == &IHostMemoryManager::IID {
            let tmp = MyHostMemoryManager {};
            let memory_manager: IHostMemoryManager = tmp.into();
            unsafe {
                memory_manager
                    .query(&*_riid, _ppobject as *mut *const c_void)
                    .ok()?
            };

            S_OK.ok()
        } else {
            unsafe { *_ppobject = null_mut() };
            E_NOINTERFACE.ok()
        }
    }

    fn SetAppDomainManager(
        &self,
        _dwappdomainid: u32,
        _punkappdomainmanager: ::core::option::Option<&::windows_core::IUnknown>,
    ) -> ::windows_core::Result<()> {
        Ok(())
    }
}

#[implement(IHostMemoryManager)]
pub struct MyHostMemoryManager;

impl IHostMemoryManager_Impl for MyHostMemoryManager {
    fn CreateMalloc(&self, dwmalloctype: u32) -> ::windows_core::Result<IHostMalloc> {
        let my_host_malloc = MyHostMalloc {
            m_hMallocHeap: unsafe { HeapCreate(HEAP_NO_SERIALIZE, 0, 0).unwrap() },
        };
        let mut tmp1: IHostMalloc = my_host_malloc.into();

        Ok(tmp1)
    }

    fn VirtualAlloc(
        &self,
        paddress: *const ::core::ffi::c_void,
        dwsize: usize,
        flallocationtype: u32,
        flprotect: u32,
        ecriticallevel: EMemoryCriticalLevel,
        ppmem: *mut *mut ::core::ffi::c_void,
    ) -> ::windows_core::Result<()> {
        unsafe {
            *ppmem = *ManuallyDrop::new(VirtualAlloc(
                Some(paddress),
                dwsize,
                VIRTUAL_ALLOCATION_TYPE(flallocationtype),
                PAGE_PROTECTION_FLAGS(flprotect),
            ))
        };

        if flallocationtype == 8192 && dwsize > 65536 {
            HASHMAP
                .lock()
                .unwrap()
                .push((unsafe { *ppmem } as u64, dwsize));
        }

        S_OK.ok()
    }

    fn VirtualFree(
        &self,
        lpaddress: *const ::core::ffi::c_void,
        dwsize: usize,
        dwfreetype: u32,
    ) -> ::windows_core::Result<()> {
        unsafe {
            VirtualFree(
                lpaddress as *mut c_void,
                dwsize,
                VIRTUAL_FREE_TYPE(dwfreetype),
            )?;
        };

        S_OK.ok()
    }

    fn VirtualQuery(
        &self,
        lpaddress: *const ::core::ffi::c_void,
        lpbuffer: *mut ::core::ffi::c_void,
        dwlength: usize,
        presult: *mut usize,
    ) -> ::windows_core::Result<()> {
        unsafe {
            *presult = VirtualQuery(
                Some(lpaddress),
                lpbuffer as *mut _ as *mut MEMORY_BASIC_INFORMATION,
                dwlength,
            );
        };

        S_OK.ok()
    }

    fn VirtualProtect(
        &self,
        lpaddress: *const ::core::ffi::c_void,
        dwsize: usize,
        flnewprotect: u32,
    ) -> ::windows_core::Result<u32> {
        if flnewprotect == 0 {
            return Ok(S_OK.0 as u32);
        }

        let mut old = PAGE_PROTECTION_FLAGS(0);
        unsafe {
            VirtualProtect(
                lpaddress,
                dwsize,
                PAGE_PROTECTION_FLAGS(flnewprotect),
                &mut old,
            )?;
        };

        Ok(S_OK.0 as u32)
    }

    fn GetMemoryLoad(
        &self,
        pmemoryload: *mut u32,
        pavailablebytes: *mut usize,
    ) -> ::windows_core::Result<()> {
        unsafe {
            *pmemoryload = 30;
            *pavailablebytes = 100 * 1024 * 1024;
        };

        S_OK.ok()
    }

    fn RegisterMemoryNotificationCallback(
        &self,
        pcallback: ::core::option::Option<&ICLRMemoryNotificationCallback>,
    ) -> ::windows_core::Result<()> {
        S_OK.ok()
    }

    fn NeedsVirtualAddressSpace(
        &self,
        startaddress: *const ::core::ffi::c_void,
        size: usize,
    ) -> ::windows_core::Result<()> {
        S_OK.ok()
    }

    fn AcquiredVirtualAddressSpace(
        &self,
        startaddress: *const ::core::ffi::c_void,
        size: usize,
    ) -> ::windows_core::Result<()> {
        S_OK.ok()
    }

    fn ReleasedVirtualAddressSpace(
        &self,
        startaddress: *const ::core::ffi::c_void,
    ) -> ::windows_core::Result<()> {
        S_OK.ok()
    }
}

#[implement(IHostMalloc)]
pub struct MyHostMalloc {
    pub m_hMallocHeap: HANDLE,
}

impl IHostMalloc_Impl for MyHostMalloc {
    fn Alloc(
        &self,
        cbsize: usize,
        ecriticallevel: EMemoryCriticalLevel,
        ppmem: *mut *mut ::core::ffi::c_void,
    ) -> ::windows_core::Result<()> {
        unsafe { *ppmem = HeapAlloc(self.m_hMallocHeap, HEAP_FLAGS(1 | 4), cbsize) };

        if (unsafe { *ppmem }).is_null() {
            return E_OUTOFMEMORY.ok();
        }

        S_OK.ok()
    }

    fn DebugAlloc(
        &self,
        cbsize: usize,
        ecriticallevel: EMemoryCriticalLevel,
        pszfilename: *const u8,
        ilineno: i32,
        ppmem: *mut *mut ::core::ffi::c_void,
    ) -> ::windows_core::Result<()> {
        self.Alloc(cbsize, ecriticallevel, ppmem)
    }

    fn Free(&self, pmem: *const ::core::ffi::c_void) -> ::windows_core::Result<()> {
        unsafe {
            HeapFree(
                self.m_hMallocHeap,
                windows::Win32::System::Memory::HEAP_FLAGS(0),
                Some(pmem),
            )?
        };

        S_OK.ok()
    }
}

lazy_static! {
    static ref HASHMAP: Mutex<Vec<(u64, usize)>> = Mutex::new(vec![]);
}

fn main() -> windows::core::Result<()> {
    let mut args: Vec<String> = env::args().collect();
    let mut assembly_contents = fs::read(args[1].clone()).expect("Unable to read file");

    let mut arguments: Vec<String> = vec![];
    arguments = args.split_off(2);

    unsafe {
        let metahost: ICLRMetaHost = CLRCreateInstance(&CLSID_CLRMetaHost)?;
        let runtime: ICLRRuntimeInfo = metahost.GetRuntime(w!("v4.0.30319"))?;
        let runtimehost: ICLRRuntimeHost = runtime.GetInterface(&CLRRuntimeHost)?;

        let mut tmp = MyHostControl {};
        let mut control: IHostControl = tmp.into();
        runtimehost.SetHostControl(&control).unwrap();
        runtimehost.Start()?;

        let runtimehost2: ICorRuntimeHost = runtime.GetInterface(&CorRuntimeHost)?;

        let default_domain: IUnknown = runtimehost2
            .GetDefaultDomain()
            .map_err(|e| format!("{}", e))
            .unwrap();

        let mut app_domain: *mut AppDomain = null_mut();

        let tmp = &**(default_domain.as_raw()
            as *mut *mut <IUnknown as windows::core::Interface>::Vtable);
        let _res = (tmp.QueryInterface)(
            default_domain.as_raw(),
            &AppDomain::IID,
            &mut app_domain as *mut *mut _ as *mut *const c_void,
        );

        let safe_array = create_assembly_safearray(&mut assembly_contents).unwrap();
        let safe_array_final = create_final_array(arguments).unwrap();
        let assembly = (*app_domain).load_assembly(safe_array).unwrap();
        let method_info = (*assembly).get_entrypoint().unwrap();

        (*method_info).invoke_assembly(safe_array_final).unwrap();

        unsafe { SafeArrayDestroy(safe_array) };
        unsafe { SafeArrayDestroy(safe_array_final) };
    }

    let test1 = HASHMAP.lock().unwrap()[2].0;
    let test2 = HASHMAP.lock().unwrap()[2].1;

    assembly_contents.zeroize();

    unsafe { encrypt_heap(test1, test2) };

    // Pause
    let mut buf = String::new();
    std::io::stdin().read_line(&mut buf).unwrap();

    Ok(())
}

unsafe fn encrypt_heap(heap: u64, heap_size: usize) {
    let mut array: Vec<MEMORY_BASIC_INFORMATION> = vec![];
    let mut start_address = heap as *const c_void;
    let mut i = 0;

    while (start_address as u64) < (heap + heap_size as u64) {
        array.push(unsafe { mem::zeroed() });
        VirtualQuery(Some(start_address), &mut array[i], 48);
        start_address = start_address.add(array.last().unwrap().RegionSize);
        i = i + 1;
    }

    let mut key_not: [i8; 16] = [
        0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66,
        0x66,
    ];

    let mut key = UString {
        length: key_not.len() as u32,
        maximum_length: key_not.len() as u32,
        buffer: key_not.as_mut_ptr() as *mut c_void,
    };

    let system_function032 = get_function_from_dll("Advapi32\0", "SystemFunction032\0").unwrap();
    let system_function032_fn: FnSystemFunction032 = unsafe { mem::transmute(system_function032) };

    for region in array {
        if region.State == MEM_RESERVE {
            continue;
        }

        let mut data = UString {
            length: region.RegionSize as u32,
            maximum_length: region.RegionSize as u32,
            buffer: region.BaseAddress,
        };

        system_function032_fn(&mut data as *mut _, &mut key as *mut _);
    }
}

pub unsafe fn get_function_from_dll(
    dll_name: &str,
    function_name: &str,
) -> Result<usize, io::Error> {
    let dll_handle = LoadLibraryA(PCSTR(String::from(dll_name).as_ptr())).unwrap();

    let func_adress = GetProcAddress(dll_handle, PCSTR(String::from(function_name).as_ptr()));
    match func_adress {
        None => Err(io::Error::last_os_error()),
        _ => Ok(func_adress.unwrap() as usize),
    }
}

fn create_assembly_safearray(
    assembly_contents: &mut Vec<u8>,
) -> core::result::Result<*mut SAFEARRAY, String> {
    let safe_array = unsafe { SafeArrayCreateVector(VT_UI1, 0, assembly_contents.len() as u32) };
    if safe_array.is_null() {
        return Err("SafeArrayCreate() got an error !".to_string());
    }

    unsafe {
        unsafe {
            ptr::copy_nonoverlapping(
                assembly_contents.as_ptr(),
                (*safe_array).pvData.cast(),
                assembly_contents.len(),
            )
        };
    };

    Ok(safe_array)
}

fn create_final_array(arguments: Vec<String>) -> Result<*mut SAFEARRAY, String> {
    let safe_array_args = unsafe { SafeArrayCreateVector(VT_BSTR, 0, arguments.len() as u32) };
    if safe_array_args.is_null() {
        return Err("SafeArrayCreate() got an error !".to_string());
    }

    for (i, args) in arguments.iter().enumerate() {
        let res = unsafe {
            SafeArrayPutElement(
                safe_array_args,
                addr_of!(i) as *const i32,
                BSTR::from(args).into_raw() as *const c_void,
            )
        };

        match res {
            Err(e) => {
                return Err(format!("SafeArrayPutElement() error: {}", e));
            }
            Ok(_) => {}
        }
    }

    let mut args_variant = VARIANT {
        Anonymous: VARIANT_0 {
            Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                vt: VARENUM(VT_BSTR.0 | VT_ARRAY.0),
                wReserved1: 0,
                wReserved2: 0,
                wReserved3: 0,
                Anonymous: VARIANT_0_0_0 {
                    parray: safe_array_args,
                },
            }),
        },
    };

    let safe_array_final = unsafe { SafeArrayCreateVector(VT_VARIANT, 0, 1_u32) };
    if safe_array_final.is_null() {
        return Err("SafeArrayCreate() got an error !".to_string());
    }

    let idx = 0;
    let res = unsafe {
        SafeArrayPutElement(
            safe_array_final,
            addr_of!(idx),
            addr_of_mut!(args_variant) as *const c_void,
        )
    };

    match res {
        Err(e) => Err(format!("SafeArrayPutElement() error: {}", e)),
        Ok(_) => Ok(safe_array_final),
    }
}

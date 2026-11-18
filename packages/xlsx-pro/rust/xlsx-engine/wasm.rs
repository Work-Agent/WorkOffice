//! WASI reactor ABI（`node:wasi` 直连，无 stdin/stdout 协议）。
//!
//! 契约（对应 `src/engine/wasm.ts`）：
//!
//! ```text
//! wasi_alloc(len) -> ptr        宿主把请求 JSON 写进这块内存
//! call(ptr, len)  -> i64        hi32 = 响应 ptr，lo32 = 响应 len（UTF-8 JSON）
//! wasi_free(ptr, len)           释放上面两种缓冲区
//! xlsx_protocol_version() -> u32  协议版本（宿主启动时断言）
//! ```
//!
//! 每次 `call` 处理**一个**请求，复用的是上游 `Engine` 在 `single-thread`
//! 下的既有语义（`Engine::handle_line`）——所以这里不需要 `main.rs` 的
//! reader 线程、也不需要 `serve_sequentially` 的 stdin 循环。
//!
//! 崩溃/取消语义：单线程下没有 reader 线程可以插队，`Cancel` 只能命中
//! **下一个**请求（当前请求跑完为止）。这是 wasm 接入的已知落差，写在
//! `RuntimeReport.degraded = ['cancel-unavailable']` 里，不做假承诺。

use std::cell::RefCell;

use crate::protocol::{Engine, PROTOCOL_VERSION};

thread_local! {
    /// 一次 `call` 一个请求，所以一个引擎实例就够；`thread_local` 让
    /// `Engine`（内部含 `Mutex`/`Arc`）不必是 `Sync`。
    static ENGINE: RefCell<Engine> = RefCell::new(Engine::new());
}

/// 把 `Vec<u8>` 的**精确长度**所有权交给宿主：`into_boxed_slice` 保证
/// 指针 + 长度可以原样还原，`wasi_free` 才能安全重建。
fn leak_exact(bytes: Vec<u8>) -> *mut u8 {
    Box::into_raw(bytes.into_boxed_slice()) as *mut u8
}

#[unsafe(no_mangle)]
pub extern "C" fn wasi_alloc(len: usize) -> *mut u8 {
    leak_exact(vec![0u8; len])
}

/// 安全性：`ptr` / `len` 必须来自本模块的 `wasi_alloc` 或 `call` 的返回值。
#[unsafe(no_mangle)]
pub unsafe extern "C" fn wasi_free(ptr: *mut u8, len: usize) {
    if ptr.is_null() || len == 0 {
        return;
    }
    drop(unsafe { Box::from_raw(core::ptr::slice_from_raw_parts_mut(ptr, len)) });
}

/// WASI reactor 入口：`node:wasi` 的 `wasi.initialize(instance)` 会调用它。
/// rustc 的 wasip1 cdylib **不会**自动导出同名符号（实测 1.99.1），所以要显式给。
/// Rust 侧没有需要在此运行的静态构造器；若上游将来引入，改成
/// `unsafe { __wasm_call_ctors() }` 即可。
#[unsafe(no_mangle)]
pub extern "C" fn _initialize() {}

fn protocol_version() -> u32 {
    PROTOCOL_VERSION as u32
}

#[unsafe(no_mangle)]
pub extern "C" fn xlsx_protocol_version() -> u32 {
    protocol_version()
}

/// 安全性：`ptr` / `len` 必须指向宿主通过 `wasi_alloc` 获得、且已写入
/// UTF-8 JSON 的可读区间。
#[unsafe(no_mangle)]
pub unsafe extern "C" fn call(ptr: *const u8, len: usize) -> u64 {
    let input = unsafe { core::slice::from_raw_parts(ptr, len) };
    let text = match core::str::from_utf8(input) {
        Ok(text) => dispatch(text),
        Err(error) => failure_json(&format!("Request is not valid UTF-8: {error}")),
    };
    respond(text.into_bytes())
}

fn dispatch(line: &str) -> String {
    // 线程可用时 recalc 会把结果写到 stdout 并返回 None；wasm 走的是
    // `single-thread`，因此这里应当是 `Some`。真遇到 None 也给结构化错误，
    // 不让宿主拿到空指针。
    ENGINE
        .with(|engine| engine.borrow_mut().handle_line(line))
        .unwrap_or_else(|| failure_json("Engine returned no response (threaded recalc path)"))
}

fn failure_json(message: &str) -> String {
    let escaped = message.replace('\\', "\\\\").replace('"', "\\\"");
    format!(
        r#"{{"version":{},"requestId":"","ok":false,"error":{{"code":"wasm_abi","message":"{escaped}"}}}}"#,
        protocol_version()
    )
}

fn respond(bytes: Vec<u8>) -> u64 {
    let len = bytes.len() as u64;
    let ptr = leak_exact(bytes) as u64;
    (ptr << 32) | len
}

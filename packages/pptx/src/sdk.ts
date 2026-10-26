// ─── WorkStudio SDK 运行时桥 ─────────────────────────────────────────
//
// 外部插件包在安装目录里运行，拿不到 WorkStudio 源码的相对路径。
// 宿主在装载插件前把 SDK 注册面挂到 globalThis；这里只做类型化转发。
//
// 注册面与 WorkStudio src/main/ura/sdk.ts 保持一致：
//   transport / codec / store / presenter / sidecar / collection / viewAction

import type {
  Codec,
  CollectionProvider,
  Presenter,
  SidecarProvider,
  Store,
  Transport,
  ViewAction,
} from './ura-types'

interface WorkStudioSdk {
  registerTransport(transport: Transport): void
  registerCodec(codec: Codec): void
  registerStore(store: Store): void
  registerPresenter(presenter: Presenter): void
  registerSidecar(provider: SidecarProvider): void
  registerCollection(provider: CollectionProvider): void
  registerViewAction(action: ViewAction): void
}

function sdk(): WorkStudioSdk {
  const value = (globalThis as { __workstudioSdk?: WorkStudioSdk })
    .__workstudioSdk
  if (!value) {
    throw new Error('WorkStudio SDK 未注入：插件必须由 WorkStudio 宿主装载')
  }
  return value
}

export const registerTransport = (transport: Transport): void =>
  sdk().registerTransport(transport)

export const registerCodec = (codec: Codec): void => sdk().registerCodec(codec)

export const registerStore = (store: Store): void => sdk().registerStore(store)

/** 表现层：页面 / 浏览器。 */
export const registerPresenter = (presenter: Presenter): void =>
  sdk().registerPresenter(presenter)

/** agent 读取面：文本 / schema。 */
export const registerSidecar = (provider: SidecarProvider): void =>
  sdk().registerSidecar(provider)

export const registerCollection = (provider: CollectionProvider): void =>
  sdk().registerCollection(provider)

export const registerViewAction = (action: ViewAction): void =>
  sdk().registerViewAction(action)

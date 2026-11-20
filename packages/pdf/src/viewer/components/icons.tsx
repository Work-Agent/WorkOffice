import type { SVGProps } from 'react'

type Props = SVGProps<SVGSVGElement>

function Svg({ children, ...props }: Props) {
  return (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      {...props}
    >
      {children}
    </svg>
  )
}

export const ChevronLeft = (props: Props) => (
  <Svg {...props}>
    <path d="m15 18-6-6 6-6" />
  </Svg>
)

export const ChevronRight = (props: Props) => (
  <Svg {...props}>
    <path d="m9 18 6-6-6-6" />
  </Svg>
)

export const ZoomInIcon = (props: Props) => (
  <Svg {...props}>
    <circle cx="11" cy="11" r="7" />
    <path d="m21 21-4.3-4.3M11 8v6M8 11h6" />
  </Svg>
)

export const ZoomOutIcon = (props: Props) => (
  <Svg {...props}>
    <circle cx="11" cy="11" r="7" />
    <path d="m21 21-4.3-4.3M8 11h6" />
  </Svg>
)

export const FitWidthIcon = (props: Props) => (
  <Svg {...props}>
    <path d="M4 5v14M20 5v14M8 12h8M8 12l2-2M8 12l2 2M16 12l-2-2M16 12l-2 2" />
  </Svg>
)

export const FitPageIcon = (props: Props) => (
  <Svg {...props}>
    <rect x="5" y="3" width="14" height="18" rx="1" />
    <path d="M9 8h6M9 12h6M9 16h4" />
  </Svg>
)

export const RotateIcon = (props: Props) => (
  <Svg {...props}>
    <path d="M21 12a9 9 0 1 1-3-6.7" />
    <path d="M21 3v6h-6" />
  </Svg>
)

export const ThumbsIcon = (props: Props) => (
  <Svg {...props}>
    <rect x="3" y="3" width="7" height="7" rx="1" />
    <rect x="14" y="3" width="7" height="7" rx="1" />
    <rect x="3" y="14" width="7" height="7" rx="1" />
    <rect x="14" y="14" width="7" height="7" rx="1" />
  </Svg>
)

export const OutlineIcon = (props: Props) => (
  <Svg {...props}>
    <path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01" />
  </Svg>
)

export const SearchIcon = (props: Props) => (
  <Svg {...props}>
    <circle cx="11" cy="11" r="7" />
    <path d="m21 21-4.3-4.3" />
  </Svg>
)

export const PrinterIcon = (props: Props) => (
  <Svg {...props}>
    <path d="M6 9V3h12v6M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" />
    <rect x="6" y="14" width="12" height="8" rx="1" />
  </Svg>
)

export const DownloadIcon = (props: Props) => (
  <Svg {...props}>
    <path d="M12 3v12M7 10l5 5 5-5M5 21h14" />
  </Svg>
)

export const EyeIcon = (props: Props) => (
  <Svg {...props}>
    <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z" />
    <circle cx="12" cy="12" r="3" />
  </Svg>
)

export const CloseIcon = (props: Props) => (
  <Svg {...props}>
    <path d="m18 6-12 12M6 6l12 12" />
  </Svg>
)

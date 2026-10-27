import React, { useId } from 'react';

const Logo = ({ className, src }: { className?: string; src?: string }) => {
  const uid = useId().replace(/:/g, '');
  const bgId = `hgBg_${uid}`;
  const goldId = `hgGold_${uid}`;

  if (src) {
    return <img src={src} alt="Logo Hoàng Gia" className={className} referrerPolicy="no-referrer" />;
  }

  return (
    <svg
      viewBox="0 0 512 512"
      className={className}
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      shapeRendering="geometricPrecision"
      aria-label="Logo Hệ thống Giáo dục Hoàng Gia"
    >
      <defs>
        <linearGradient id={bgId} x1="0" y1="0" x2="512" y2="512" gradientUnits="userSpaceOnUse">
          <stop offset="0%" stopColor="#059669" />
          <stop offset="55%" stopColor="#0D9488" />
          <stop offset="100%" stopColor="#0F766E" />
        </linearGradient>
        <linearGradient id={goldId} x1="120" y1="90" x2="392" y2="410" gradientUnits="userSpaceOnUse">
          <stop offset="0%" stopColor="#FEF08A" />
          <stop offset="48%" stopColor="#FBBF24" />
          <stop offset="100%" stopColor="#D97706" />
        </linearGradient>
      </defs>

      {/* Crisp Rounded Emblem Base */}
      <rect x="16" y="16" width="480" height="480" rx="116" fill={`url(#${bgId})`} />
      <rect
        x="36"
        y="36"
        width="440"
        height="440"
        rx="98"
        stroke="#FFFFFF"
        strokeOpacity="0.26"
        strokeWidth="6"
      />

      {/* Royal Academic Crown Crest */}
      <path
        d="M192 126L225 154L256 112L287 154L320 126L309 172H203L192 126Z"
        fill={`url(#${goldId})`}
      />
      <circle cx="256" cy="98" r="10" fill="#FEF08A" />

      {/* Sharp Geometric 'H' */}
      <path
        d="M118 196H162V274H224V196H268V386H224V312H162V386H118V196Z"
        fill="#FFFFFF"
      />

      {/* Sharp Geometric 'G' */}
      <path
        d="M364 192C396 192 420 208 432 236L392 253C385 237 374 230 359 230C332 230 314 253 314 291C314 329 332 351 362 351C383 351 395 340 398 318H358V283H438V386H404L400 365C389 380 373 389 351 389C303 389 270 349 270 291C270 231 306 192 364 192Z"
        fill={`url(#${goldId})`}
      />

      {/* Open Knowledge Book Flourish */}
      <path
        d="M132 422C178 407 222 412 256 428C290 412 334 407 380 422"
        stroke="#FFFFFF"
        strokeOpacity="0.92"
        strokeWidth="14"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
};

export default Logo;

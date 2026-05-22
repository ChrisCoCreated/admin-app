export const MAX_SELECTION = 6;
export const PAN_LIMIT = 100;
export const EXPORT_WIDTH = 2000;
export const DEFAULT_GAP_PX = 24;
export const DEFAULT_RADIUS_PX = 36;

export const LAYOUT_BACKGROUND_COLORS = new Set([
  "#ffffff",
  "#000000",
  "#1db9d2",
  "#6d3ac6",
  "#ce3f8f",
  "#3aceac",
  "#ffd23f",
]);

export const LAYOUTS = [
  {
    id: "single",
    name: "Single",
    aspect: 1.25,
    slots: [{ x: 0, y: 0, w: 1, h: 1 }],
  },
  {
    id: "split_two",
    name: "Split Two",
    aspect: 1.5,
    slots: [
      { x: 0, y: 0, w: 0.5, h: 1 },
      { x: 0.5, y: 0, w: 0.5, h: 1 },
    ],
  },
  {
    id: "stacked_two",
    name: "Stacked Two",
    aspect: 1.25,
    slots: [
      { x: 0, y: 0, w: 1, h: 0.5 },
      { x: 0, y: 0.5, w: 1, h: 0.5 },
    ],
  },
  {
    id: "feature_two",
    name: "Feature Two",
    aspect: 1.5,
    slots: [
      { x: 0, y: 0, w: 0.62, h: 1 },
      { x: 0.62, y: 0, w: 0.38, h: 1 },
    ],
  },
  {
    id: "feature_three",
    name: "Feature Three",
    aspect: 1.3333,
    slots: [
      { x: 0, y: 0, w: 1, h: 0.55 },
      { x: 0, y: 0.55, w: 0.5, h: 0.45 },
      { x: 0.5, y: 0.55, w: 0.5, h: 0.45 },
    ],
  },
  {
    id: "columns_three",
    name: "Columns Three",
    aspect: 1.5,
    slots: [
      { x: 0, y: 0, w: 1 / 3, h: 1 },
      { x: 1 / 3, y: 0, w: 1 / 3, h: 1 },
      { x: 2 / 3, y: 0, w: 1 / 3, h: 1 },
    ],
  },
  {
    id: "side_stack_three",
    name: "Side Stack Three",
    aspect: 1.4,
    slots: [
      { x: 0, y: 0, w: 0.58, h: 1 },
      { x: 0.58, y: 0, w: 0.42, h: 0.5 },
      { x: 0.58, y: 0.5, w: 0.42, h: 0.5 },
    ],
  },
  {
    id: "grid_four",
    name: "Grid Four",
    aspect: 1,
    slots: [
      { x: 0, y: 0, w: 0.5, h: 0.5 },
      { x: 0.5, y: 0, w: 0.5, h: 0.5 },
      { x: 0, y: 0.5, w: 0.5, h: 0.5 },
      { x: 0.5, y: 0.5, w: 0.5, h: 0.5 },
    ],
  },
  {
    id: "portrait_four",
    name: "Portrait Four",
    aspect: 0.75,
    slots: [
      { x: 0, y: 0, w: 0.5, h: 0.5 },
      { x: 0.5, y: 0, w: 0.5, h: 0.5 },
      { x: 0, y: 0.5, w: 0.5, h: 0.5 },
      { x: 0.5, y: 0.5, w: 0.5, h: 0.5 },
    ],
  },
  {
    id: "landscape_four",
    name: "Landscape Four",
    aspect: 1.3333,
    slots: [
      { x: 0, y: 0, w: 0.5, h: 0.5 },
      { x: 0.5, y: 0, w: 0.5, h: 0.5 },
      { x: 0, y: 0.5, w: 0.5, h: 0.5 },
      { x: 0.5, y: 0.5, w: 0.5, h: 0.5 },
    ],
  },
  {
    id: "feature_four_left",
    name: "Feature Four Left",
    aspect: 1.4,
    slots: [
      { x: 0, y: 0, w: 0.58, h: 1 },
      { x: 0.58, y: 0, w: 0.42, h: 0.4 },
      { x: 0.58, y: 0.4, w: 0.21, h: 0.6 },
      { x: 0.79, y: 0.4, w: 0.21, h: 0.6 },
    ],
  },
  {
    id: "banner_four",
    name: "Banner Four",
    aspect: 1.45,
    slots: [
      { x: 0, y: 0, w: 1, h: 0.58 },
      { x: 0, y: 0.58, w: 1 / 3, h: 0.42 },
      { x: 1 / 3, y: 0.58, w: 1 / 3, h: 0.42 },
      { x: 2 / 3, y: 0.58, w: 1 / 3, h: 0.42 },
    ],
  },
  {
    id: "mosaic_five",
    name: "Mosaic Five",
    aspect: 1.3333,
    slots: [
      { x: 0, y: 0, w: 0.6, h: 0.62 },
      { x: 0.6, y: 0, w: 0.4, h: 0.31 },
      { x: 0.6, y: 0.31, w: 0.4, h: 0.31 },
      { x: 0, y: 0.62, w: 0.5, h: 0.38 },
      { x: 0.5, y: 0.62, w: 0.5, h: 0.38 },
    ],
  },
  {
    id: "strip_five",
    name: "Strip Five",
    aspect: 1.6,
    slots: [
      { x: 0, y: 0, w: 1 / 3, h: 0.55 },
      { x: 1 / 3, y: 0, w: 1 / 3, h: 0.55 },
      { x: 2 / 3, y: 0, w: 1 / 3, h: 0.55 },
      { x: 0, y: 0.55, w: 0.5, h: 0.45 },
      { x: 0.5, y: 0.55, w: 0.5, h: 0.45 },
    ],
  },
  {
    id: "feature_five_right",
    name: "Feature Five Right",
    aspect: 1.45,
    slots: [
      { x: 0, y: 0, w: 0.35, h: 0.5 },
      { x: 0, y: 0.5, w: 0.35, h: 0.5 },
      { x: 0.35, y: 0, w: 0.65, h: 0.64 },
      { x: 0.35, y: 0.64, w: 0.325, h: 0.36 },
      { x: 0.675, y: 0.64, w: 0.325, h: 0.36 },
    ],
  },
  {
    id: "grid_six",
    name: "Grid Six",
    aspect: 1.5,
    slots: [
      { x: 0, y: 0, w: 1 / 3, h: 0.5 },
      { x: 1 / 3, y: 0, w: 1 / 3, h: 0.5 },
      { x: 2 / 3, y: 0, w: 1 / 3, h: 0.5 },
      { x: 0, y: 0.5, w: 1 / 3, h: 0.5 },
      { x: 1 / 3, y: 0.5, w: 1 / 3, h: 0.5 },
      { x: 2 / 3, y: 0.5, w: 1 / 3, h: 0.5 },
    ],
  },
  {
    id: "portrait_six",
    name: "Portrait Six",
    aspect: 1.2,
    slots: [
      { x: 0, y: 0, w: 1 / 3, h: 0.5 },
      { x: 1 / 3, y: 0, w: 1 / 3, h: 0.5 },
      { x: 2 / 3, y: 0, w: 1 / 3, h: 0.5 },
      { x: 0, y: 0.5, w: 1 / 3, h: 0.5 },
      { x: 1 / 3, y: 0.5, w: 1 / 3, h: 0.5 },
      { x: 2 / 3, y: 0.5, w: 1 / 3, h: 0.5 },
    ],
  },
  {
    id: "columns_six",
    name: "Columns Six",
    aspect: 1.35,
    slots: [
      { x: 0, y: 0, w: 0.5, h: 1 / 3 },
      { x: 0, y: 1 / 3, w: 0.5, h: 1 / 3 },
      { x: 0, y: 2 / 3, w: 0.5, h: 1 / 3 },
      { x: 0.5, y: 0, w: 0.5, h: 1 / 3 },
      { x: 0.5, y: 1 / 3, w: 0.5, h: 1 / 3 },
      { x: 0.5, y: 2 / 3, w: 0.5, h: 1 / 3 },
    ],
  },
  {
    id: "feature_six",
    name: "Feature Six",
    aspect: 1.5,
    slots: [
      { x: 0, y: 0, w: 0.5, h: 0.66 },
      { x: 0.5, y: 0, w: 0.25, h: 0.33 },
      { x: 0.75, y: 0, w: 0.25, h: 0.33 },
      { x: 0.5, y: 0.33, w: 0.5, h: 0.33 },
      { x: 0, y: 0.66, w: 1 / 3, h: 0.34 },
      { x: 1 / 3, y: 0.66, w: 2 / 3, h: 0.34 },
    ],
  },
];

export function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

export function computeLogoOverlayRect(width, height, gapPx, image) {
  const safeWidth = Math.max(1, Number(width) || 1);
  const safeHeight = Math.max(1, Number(height) || 1);
  const imageWidth = Math.max(1, Number(image?.width) || 1);
  const imageHeight = Math.max(1, Number(image?.height) || 1);
  const margin = Math.max(16, Number(gapPx) || 0);
  const maxWidth = safeWidth * 0.26;
  const maxHeight = safeHeight * 0.18;
  const scale = Math.min(maxWidth / imageWidth, maxHeight / imageHeight);
  const drawW = Math.max(40, imageWidth * scale);
  const drawH = Math.max(24, imageHeight * scale);
  return {
    x: safeWidth - margin - drawW,
    y: safeHeight - margin - drawH,
    w: drawW,
    h: drawH,
  };
}

export function computeImagePlacement(slotWidth, slotHeight, imageWidth, imageHeight, zoom, panX, panY, targetAspect = null) {
  const safeSlotWidth = Math.max(1, Number(slotWidth) || 1);
  const safeSlotHeight = Math.max(1, Number(slotHeight) || 1);
  const safeImageWidth = Math.max(1, Number(imageWidth) || 1);
  const safeImageHeight = Math.max(1, Number(imageHeight) || 1);
  const safeZoom = clamp(Number(zoom) || 1, 1, 3);
  const safePanX = clamp(Number(panX) || 0, -PAN_LIMIT, PAN_LIMIT);
  const safePanY = clamp(Number(panY) || 0, -PAN_LIMIT, PAN_LIMIT);

  let sourceX = 0;
  let sourceY = 0;
  let sourceW = safeImageWidth;
  let sourceH = safeImageHeight;
  if (targetAspect && Number.isFinite(targetAspect) && targetAspect > 0) {
    const sourceAspect = safeImageWidth / safeImageHeight;
    if (sourceAspect > targetAspect) {
      sourceW = safeImageHeight * targetAspect;
      sourceX = (safeImageWidth - sourceW) * 0.5;
    } else if (sourceAspect < targetAspect) {
      sourceH = safeImageWidth / targetAspect;
      sourceY = (safeImageHeight - sourceH) * 0.5;
    }
  }

  const baseScale = Math.max(safeSlotWidth / sourceW, safeSlotHeight / sourceH);
  const scale = baseScale * safeZoom;
  const drawW = sourceW * scale;
  const drawH = sourceH * scale;
  const overflowX = Math.max(0, drawW - safeSlotWidth);
  const overflowY = Math.max(0, drawH - safeSlotHeight);
  const maxOffsetX = overflowX * 0.5;
  const maxOffsetY = overflowY * 0.5;
  const offsetX = (safePanX / PAN_LIMIT) * maxOffsetX;
  const offsetY = (safePanY / PAN_LIMIT) * maxOffsetY;
  const drawX = (safeSlotWidth - drawW) * 0.5 + offsetX;
  const drawY = (safeSlotHeight - drawH) * 0.5 + offsetY;

  return {
    sourceX,
    sourceY,
    sourceW,
    sourceH,
    scale,
    drawX,
    drawY,
    drawW,
    drawH,
    renderX: drawX - sourceX * scale,
    renderY: drawY - sourceY * scale,
    renderW: safeImageWidth * scale,
    renderH: safeImageHeight * scale,
    maxOffsetX,
    maxOffsetY,
  };
}

function rangesOverlap(startA, endA, startB, endB) {
  return Math.min(endA, endB) - Math.max(startA, startB) > 1e-6;
}

function nearlyEqual(valueA, valueB) {
  return Math.abs(valueA - valueB) <= 1e-6;
}

export function getSlotNeighborFlags(slots, index) {
  const slot = slots[index];
  const slotLeft = slot.x;
  const slotRight = slot.x + slot.w;
  const slotTop = slot.y;
  const slotBottom = slot.y + slot.h;
  const flags = {
    hasLeftNeighbor: false,
    hasRightNeighbor: false,
    hasTopNeighbor: false,
    hasBottomNeighbor: false,
  };

  for (let otherIndex = 0; otherIndex < slots.length; otherIndex += 1) {
    if (otherIndex === index) {
      continue;
    }
    const other = slots[otherIndex];
    const otherLeft = other.x;
    const otherRight = other.x + other.w;
    const otherTop = other.y;
    const otherBottom = other.y + other.h;

    if (nearlyEqual(slotLeft, otherRight) && rangesOverlap(slotTop, slotBottom, otherTop, otherBottom)) {
      flags.hasLeftNeighbor = true;
    }
    if (nearlyEqual(slotRight, otherLeft) && rangesOverlap(slotTop, slotBottom, otherTop, otherBottom)) {
      flags.hasRightNeighbor = true;
    }
    if (nearlyEqual(slotTop, otherBottom) && rangesOverlap(slotLeft, slotRight, otherLeft, otherRight)) {
      flags.hasTopNeighbor = true;
    }
    if (nearlyEqual(slotBottom, otherTop) && rangesOverlap(slotLeft, slotRight, otherLeft, otherRight)) {
      flags.hasBottomNeighbor = true;
    }
  }

  return flags;
}

export function computeAdjustedSlotRect(slot, neighborFlags, width, height, gapPx) {
  const halfGap = Math.max(0, Number(gapPx) || 0) * 0.5;
  let x = slot.x * width;
  let y = slot.y * height;
  let w = slot.w * width;
  let h = slot.h * height;

  if (neighborFlags.hasLeftNeighbor) {
    x += halfGap;
    w -= halfGap;
  }
  if (neighborFlags.hasRightNeighbor) {
    w -= halfGap;
  }
  if (neighborFlags.hasTopNeighbor) {
    y += halfGap;
    h -= halfGap;
  }
  if (neighborFlags.hasBottomNeighbor) {
    h -= halfGap;
  }

  return {
    x,
    y,
    w: Math.max(0, w),
    h: Math.max(0, h),
  };
}

export function drawRoundedRectPath(ctx, x, y, w, h, radius) {
  const maxRadius = Math.min(w, h) * 0.5;
  const r = clamp(Number(radius) || 0, 0, maxRadius);
  if (!r) {
    ctx.rect(x, y, w, h);
    return;
  }
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.arcTo(x + w, y, x + w, y + r, r);
  ctx.lineTo(x + w, y + h - r);
  ctx.arcTo(x + w, y + h, x + w - r, y + h, r);
  ctx.lineTo(x + r, y + h);
  ctx.arcTo(x, y + h, x, y + h - r, r);
  ctx.lineTo(x, y + r);
  ctx.arcTo(x, y, x + r, y, r);
}

export function drawImageIntoSlot(ctx, img, slot, transform, cornerRadiusPx = 0, targetAspect = null) {
  const slotX = slot.x;
  const slotY = slot.y;
  const slotW = slot.w;
  const slotH = slot.h;
  const zoom = clamp(Number(transform.zoom) || 1, 1, 3);
  const panX = clamp(Number(transform.panX) || 0, -PAN_LIMIT, PAN_LIMIT);
  const panY = clamp(Number(transform.panY) || 0, -PAN_LIMIT, PAN_LIMIT);

  const placement = computeImagePlacement(slotW, slotH, img.width, img.height, zoom, panX, panY, targetAspect);

  ctx.save();
  ctx.beginPath();
  drawRoundedRectPath(ctx, slotX, slotY, slotW, slotH, cornerRadiusPx);
  ctx.clip();
  ctx.drawImage(
    img,
    placement.sourceX,
    placement.sourceY,
    placement.sourceW,
    placement.sourceH,
    slotX + placement.drawX,
    slotY + placement.drawY,
    placement.drawW,
    placement.drawH
  );
  ctx.restore();
}

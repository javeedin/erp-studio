import React, { useRef, useState, useEffect } from 'react';
import { Modal, Button } from 'antd';

export interface SelectedRegion {
  x: number; y: number; w: number; h: number;
  naturalW: number; naturalH: number;
}

interface Props {
  screenshot: string;
  onConfirm: (region: SelectedRegion) => void;
  onCancel: () => void;
}

const AreaSelector: React.FC<Props> = ({ screenshot, onConfirm, onCancel }) => {
  const imgRef    = useRef<HTMLImageElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);

  const [loaded, setLoaded]       = useState(false);
  const [dispSize, setDispSize]   = useState({ w: 0, h: 0 });
  const [naturalSize, setNaturalSize] = useState({ w: 0, h: 0 });
  const [selection, setSelection] = useState<{ x: number; y: number; w: number; h: number } | null>(null);

  const isDragging = useRef(false);
  const startPos   = useRef({ x: 0, y: 0 });

  const handleImgLoad = (e: React.SyntheticEvent<HTMLImageElement>) => {
    const img  = e.currentTarget;
    const maxW = Math.min(window.innerWidth * 0.85, 1200);
    const scale = img.naturalWidth > maxW ? maxW / img.naturalWidth : 1;
    setNaturalSize({ w: img.naturalWidth, h: img.naturalHeight });
    setDispSize({ w: Math.round(img.naturalWidth * scale), h: Math.round(img.naturalHeight * scale) });
    setLoaded(true);
  };

  useEffect(() => {
    if (!loaded || !canvasRef.current || dispSize.w === 0) return;
    const canvas = canvasRef.current;
    canvas.width  = dispSize.w;
    canvas.height = dispSize.h;
  }, [loaded, dispSize]);

  const getPos = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current!;
    const rect   = canvas.getBoundingClientRect();
    return {
      x: Math.round((e.clientX - rect.left) * (canvas.width  / rect.width)),
      y: Math.round((e.clientY - rect.top)  * (canvas.height / rect.height)),
    };
  };

  const drawOverlay = (sel: { x: number; y: number; w: number; h: number } | null) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d')!;
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Dark overlay over everything
    ctx.fillStyle = 'rgba(0,0,0,0.45)';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    if (sel && (Math.abs(sel.w) > 2 || Math.abs(sel.h) > 2)) {
      const x = sel.w < 0 ? sel.x + sel.w : sel.x;
      const y = sel.h < 0 ? sel.y + sel.h : sel.y;
      const w = Math.abs(sel.w);
      const h = Math.abs(sel.h);

      // Cut out selected region so screenshot shows through clearly
      ctx.clearRect(x, y, w, h);

      // Blue selection border
      ctx.strokeStyle = '#2196f3';
      ctx.lineWidth   = 2;
      ctx.strokeRect(x, y, w, h);

      // White corner handles
      ctx.strokeStyle = '#fff';
      ctx.lineWidth   = 3;
      const c = 12;
      const corners: [number, number, number, number, number, number][] = [
        [x, y,   x+c, y,   x,   y+c],
        [x+w, y, x+w-c, y, x+w, y+c],
        [x, y+h, x+c, y+h, x,   y+h-c],
        [x+w, y+h, x+w-c, y+h, x+w, y+h-c],
      ];
      corners.forEach(([ax, ay, bx, by, cx, cy]) => {
        ctx.beginPath();
        ctx.moveTo(bx, by); ctx.lineTo(ax, ay); ctx.lineTo(cx, cy);
        ctx.stroke();
      });

      // Size label
      ctx.fillStyle = '#2196f3';
      ctx.fillRect(x, y > 20 ? y - 22 : y + h + 2, 80, 18);
      ctx.fillStyle = '#fff';
      ctx.font = '11px monospace';
      ctx.fillText(`${w} × ${h}`, x + 4, y > 20 ? y - 8 : y + h + 14);
    }
  };

  const onMouseDown = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const pos = getPos(e);
    isDragging.current = true;
    startPos.current   = pos;
    setSelection(null);
    drawOverlay(null);
  };

  const onMouseMove = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isDragging.current) return;
    const pos = getPos(e);
    const raw = { x: startPos.current.x, y: startPos.current.y, w: pos.x - startPos.current.x, h: pos.y - startPos.current.y };
    drawOverlay(raw);
  };

  const onMouseUp = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isDragging.current) return;
    isDragging.current = false;
    const pos = getPos(e);
    const raw = { x: startPos.current.x, y: startPos.current.y, w: pos.x - startPos.current.x, h: pos.y - startPos.current.y };
    const normalized = {
      x: raw.w < 0 ? raw.x + raw.w : raw.x,
      y: raw.h < 0 ? raw.y + raw.h : raw.y,
      w: Math.abs(raw.w),
      h: Math.abs(raw.h),
    };
    setSelection(normalized);
    drawOverlay(normalized);
  };

  const handleConfirm = () => {
    if (!selection || selection.w < 5 || selection.h < 5) return;
    const canvas = canvasRef.current!;
    const scaleX  = naturalSize.w / canvas.width;
    const scaleY  = naturalSize.h / canvas.height;
    onConfirm({
      x: Math.round(selection.x * scaleX),
      y: Math.round(selection.y * scaleY),
      w: Math.round(selection.w * scaleX),
      h: Math.round(selection.h * scaleY),
      naturalW: naturalSize.w,
      naturalH: naturalSize.h,
    });
  };

  const hasSelection = !!selection && selection.w >= 5 && selection.h >= 5;

  return (
    <Modal
      open
      onCancel={onCancel}
      width={Math.max(Math.min(dispSize.w + 80, window.innerWidth - 40), 600)}
      style={{ top: 16 }}
      destroyOnClose
      title="Select Area to Capture"
      footer={
        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
          <Button onClick={onCancel}>Cancel</Button>
          <Button
            type="primary"
            onClick={handleConfirm}
            disabled={!hasSelection}
            style={{ background: '#1565c0', borderColor: '#1565c0' }}
          >
            {hasSelection ? `Capture Area (${selection!.w} × ${selection!.h})` : 'Draw a selection first'}
          </Button>
        </div>
      }
    >
      <p style={{ color: '#888', fontSize: 12, marginBottom: 10 }}>
        Click and drag to select the region you want to capture. Only fields inside the selection will be extracted.
      </p>
      <div style={{ overflow: 'auto', maxHeight: 'calc(100vh - 270px)', background: '#222', borderRadius: 6 }}>
        <div style={{ position: 'relative', display: 'inline-block', lineHeight: 0 }}>
          <img
            ref={imgRef}
            src={screenshot}
            onLoad={handleImgLoad}
            style={{ display: 'block', maxWidth: '100%' }}
            alt="page"
            draggable={false}
          />
          {loaded && (
            <canvas
              ref={canvasRef}
              style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', cursor: 'crosshair' }}
              onMouseDown={onMouseDown}
              onMouseMove={onMouseMove}
              onMouseUp={onMouseUp}
              onMouseLeave={onMouseUp}
            />
          )}
        </div>
      </div>
    </Modal>
  );
};

export default AreaSelector;

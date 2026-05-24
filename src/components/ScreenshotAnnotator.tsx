import React, { useRef, useState, useEffect } from 'react';
import { Modal, Button, Tooltip } from 'antd';
import { UndoOutlined, SaveOutlined } from '@ant-design/icons';

type Tool = 'pen' | 'arrow' | 'rect' | 'circle';

interface Props {
  screenshot: string;
  onSave: (annotated: string) => void;
  onClose: () => void;
}

const COLORS = [
  { value: '#ff3333', label: 'Red' },
  { value: '#ff9900', label: 'Orange' },
  { value: '#ffff00', label: 'Yellow' },
  { value: '#00cc44', label: 'Green' },
  { value: '#2196f3', label: 'Blue' },
  { value: '#ffffff', label: 'White' },
  { value: '#111111', label: 'Black' },
];

const TOOLS: { key: Tool; label: string }[] = [
  { key: 'pen',    label: '✏ Pen' },
  { key: 'arrow',  label: '↗ Arrow' },
  { key: 'rect',   label: '▭ Box' },
  { key: 'circle', label: '○ Circle' },
];

const WIDTHS = [2, 4, 8];

const ScreenshotAnnotator: React.FC<Props> = ({ screenshot, onSave, onClose }) => {
  const canvasRef               = useRef<HTMLCanvasElement>(null);
  const [tool, setTool]         = useState<Tool>('arrow');
  const [color, setColor]       = useState('#ff3333');
  const [lineWidth, setLineWidth] = useState(3);
  const [undoStack, setUndoStack] = useState<ImageData[]>([]);
  const [dims, setDims]         = useState({ w: 860, h: 500 });

  const isDrawing = useRef(false);
  const startPos  = useRef<{ x: number; y: number } | null>(null);
  const baseSnap  = useRef<ImageData | null>(null);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d')!;
    const img = new Image();
    img.onload = () => {
      const maxW = 860;
      const scale = img.width > maxW ? maxW / img.width : 1;
      const w = Math.round(img.width * scale);
      const h = Math.round(img.height * scale);
      canvas.width  = w;
      canvas.height = h;
      setDims({ w, h });
      ctx.drawImage(img, 0, 0, w, h);
      baseSnap.current = ctx.getImageData(0, 0, w, h);
    };
    img.src = screenshot;
  }, [screenshot]);

  const getPos = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current!;
    const rect = canvas.getBoundingClientRect();
    return {
      x: (e.clientX - rect.left) * (canvas.width  / rect.width),
      y: (e.clientY - rect.top)  * (canvas.height / rect.height),
    };
  };

  const applyCtx = (ctx: CanvasRenderingContext2D) => {
    ctx.strokeStyle = color;
    ctx.fillStyle   = color;
    ctx.lineWidth   = lineWidth;
    ctx.lineCap     = 'round';
    ctx.lineJoin    = 'round';
  };

  const drawArrow = (
    ctx: CanvasRenderingContext2D,
    x1: number, y1: number,
    x2: number, y2: number,
  ) => {
    const headLen = Math.max(14, lineWidth * 4);
    const angle   = Math.atan2(y2 - y1, x2 - x1);
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(
      x2 - headLen * Math.cos(angle - Math.PI / 6),
      y2 - headLen * Math.sin(angle - Math.PI / 6),
    );
    ctx.lineTo(
      x2 - headLen * Math.cos(angle + Math.PI / 6),
      y2 - headLen * Math.sin(angle + Math.PI / 6),
    );
    ctx.closePath();
    ctx.fill();
  };

  const onMouseDown = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d')!;
    const pos = getPos(e);

    // snapshot before this stroke (for undo)
    const snap = ctx.getImageData(0, 0, canvas.width, canvas.height);
    setUndoStack(prev => [...prev, snap]);
    baseSnap.current = snap;

    isDrawing.current = true;
    startPos.current  = pos;
    applyCtx(ctx);

    if (tool === 'pen') {
      ctx.beginPath();
      ctx.moveTo(pos.x, pos.y);
    }
  };

  const onMouseMove = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isDrawing.current || !startPos.current) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d')!;
    const pos = getPos(e);
    applyCtx(ctx);

    if (tool === 'pen') {
      ctx.lineTo(pos.x, pos.y);
      ctx.stroke();
      return;
    }

    // shape tools: restore base then draw preview
    if (baseSnap.current) ctx.putImageData(baseSnap.current, 0, 0);
    const { x: x1, y: y1 } = startPos.current;
    const { x: x2, y: y2 } = pos;

    if (tool === 'arrow') {
      drawArrow(ctx, x1, y1, x2, y2);
    } else if (tool === 'rect') {
      ctx.beginPath();
      ctx.strokeRect(x1, y1, x2 - x1, y2 - y1);
    } else if (tool === 'circle') {
      const rx = (x2 - x1) / 2;
      const ry = (y2 - y1) / 2;
      ctx.beginPath();
      ctx.ellipse(x1 + rx, y1 + ry, Math.abs(rx), Math.abs(ry), 0, 0, Math.PI * 2);
      ctx.stroke();
    }
  };

  const onMouseUp = () => {
    if (!isDrawing.current) return;
    isDrawing.current = false;
    startPos.current  = null;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d')!;
    if (tool === 'pen') ctx.closePath();
    baseSnap.current = ctx.getImageData(0, 0, canvas.width, canvas.height);
  };

  const handleUndo = () => {
    const canvas = canvasRef.current;
    if (!canvas || undoStack.length === 0) return;
    const ctx  = canvas.getContext('2d')!;
    const prev = undoStack[undoStack.length - 1];
    ctx.putImageData(prev, 0, 0);
    baseSnap.current = prev;
    setUndoStack(s => s.slice(0, -1));
  };

  const handleSave = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    onSave(canvas.toDataURL('image/png'));
  };

  const modalWidth = Math.min(dims.w + 48, typeof window !== 'undefined' ? window.innerWidth - 40 : 960);

  const toolbar = (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', paddingRight: 32 }}>
      {/* Tools */}
      <div style={{ display: 'flex', gap: 4 }}>
        {TOOLS.map(t => (
          <Button
            key={t.key}
            size="small"
            type={tool === t.key ? 'primary' : 'default'}
            onClick={() => setTool(t.key)}
            style={tool === t.key ? { background: '#1565c0', borderColor: '#1565c0' } : {}}
          >
            {t.label}
          </Button>
        ))}
      </div>

      <div style={{ width: 1, height: 22, background: '#e0e0e0' }} />

      {/* Colors */}
      <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
        {COLORS.map(c => (
          <Tooltip key={c.value} title={c.label} placement="bottom">
            <div
              onClick={() => setColor(c.value)}
              style={{
                width: 20, height: 20, borderRadius: '50%',
                background: c.value,
                border: color === c.value
                  ? '3px solid #1890ff'
                  : c.value === '#ffffff' ? '2px solid #ccc' : '2px solid transparent',
                cursor: 'pointer',
                flexShrink: 0,
                boxSizing: 'border-box',
              }}
            />
          </Tooltip>
        ))}
      </div>

      <div style={{ width: 1, height: 22, background: '#e0e0e0' }} />

      {/* Line widths */}
      <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
        {WIDTHS.map(w => (
          <Tooltip key={w} title={`Size ${w}`} placement="bottom">
            <div
              onClick={() => setLineWidth(w)}
              style={{
                width: 30, height: 22,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                cursor: 'pointer',
                border: lineWidth === w ? '1px solid #1890ff' : '1px solid #d9d9d9',
                borderRadius: 3,
                background: lineWidth === w ? '#e6f4ff' : '#fafafa',
              }}
            >
              <div style={{ width: 16, height: w, background: '#333', borderRadius: w }} />
            </div>
          </Tooltip>
        ))}
      </div>

      <div style={{ width: 1, height: 22, background: '#e0e0e0' }} />

      <Tooltip title="Undo last mark" placement="bottom">
        <Button
          size="small"
          icon={<UndoOutlined />}
          onClick={handleUndo}
          disabled={undoStack.length === 0}
        />
      </Tooltip>
    </div>
  );

  return (
    <Modal
      open
      onCancel={onClose}
      width={modalWidth}
      centered
      destroyOnClose
      title={toolbar}
      footer={
        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
          <Button onClick={onClose}>Cancel</Button>
          <Button
            type="primary"
            icon={<SaveOutlined />}
            onClick={handleSave}
            style={{ background: '#1565c0', borderColor: '#1565c0' }}
          >
            Save Annotation
          </Button>
        </div>
      }
    >
      <div style={{ overflow: 'auto', maxHeight: 'calc(80vh - 140px)', textAlign: 'center', background: '#222' }}>
        <canvas
          ref={canvasRef}
          style={{ cursor: 'crosshair', maxWidth: '100%', display: 'block', margin: '0 auto' }}
          onMouseDown={onMouseDown}
          onMouseMove={onMouseMove}
          onMouseUp={onMouseUp}
          onMouseLeave={onMouseUp}
        />
      </div>
    </Modal>
  );
};

export default ScreenshotAnnotator;

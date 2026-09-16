import { type ReactNode, useEffect, useRef, useState } from 'react';
import { Link } from 'react-router-dom';

const REPO = 'https://github.com/b32n/spindle';

/**
 * Full-page frame for the live demos: a thin header (home · switch surface ·
 * docs · GitHub) over a measured stage. The editors need explicit pixel
 * width/height, so the stage is measured and handed to `children` as a size.
 */
export default function DemoChrome({
  active,
  hint,
  children,
}: {
  active: 'sheets' | 'docs' | 'slides';
  hint: string;
  children: (size: { width: number; height: number }) => ReactNode;
}) {
  const stageRef = useRef<HTMLDivElement>(null);
  const [size, setSize] = useState({ width: 0, height: 0 });

  useEffect(() => {
    const el = stageRef.current;
    if (!el) return;
    const measure = () => setSize({ width: el.clientWidth, height: el.clientHeight });
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    measure();
    return () => ro.disconnect();
  }, []);

  return (
    <div className="demo-shell">
      <header className="demo-top">
        <div className="demo-top-in">
          <Link className="brand" to="/" aria-label="Spindle home">
            <img src="/spindle.svg" width="24" height="24" alt="" aria-hidden="true" />
            Spindle <small>Live demo</small>
          </Link>
          <div className="demo-switch" role="tablist">
            <Link to="/demo/sheets" className={active === 'sheets' ? 'active' : ''}>
              Sheets
            </Link>
            <Link to="/demo/docs" className={active === 'docs' ? 'active' : ''}>
              Docs
            </Link>
            <Link to="/demo/slides" className={active === 'slides' ? 'active' : ''}>
              Slides
            </Link>
          </div>
          <div className="demo-top-right">
            <span className="demo-hint">{hint}</span>
            <Link to="/docs">Docs</Link>
            <a href={REPO} target="_blank" rel="noopener noreferrer">
              GitHub ↗
            </a>
          </div>
        </div>
      </header>
      <div className="demo-stage" ref={stageRef}>
        {size.width > 0 && size.height > 0 && children(size)}
      </div>
    </div>
  );
}

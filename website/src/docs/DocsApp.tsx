import { useEffect, useRef, useState } from 'react';
import { Link, useLocation, useNavigate } from 'react-router-dom';
import { docLoaders, NAV, ORDER, titleFor } from './manifest';
import { renderDoc, type TocItem } from './markdown';
import { renderMermaid } from './mermaid';

const REPO = 'https://github.com/b32n/spindle';

type Rendered = { html: string; toc: TocItem[]; hasMermaid: boolean };

// Turn /docs, /docs/, /docs/sheets/overview into the bare slug.
const slugFromPath = (pathname: string): string =>
  pathname.replace(/^\/docs\/?/, '').replace(/\/$/, '');

export default function DocsApp() {
  const location = useLocation();
  const nav = useNavigate();
  const slug = slugFromPath(location.pathname);

  const [doc, setDoc] = useState<Rendered | null>(null);
  const [missing, setMissing] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const articleRef = useRef<HTMLElement>(null);

  useEffect(() => {
    let alive = true;
    setDoc(null);
    setMissing(false);
    const loader = docLoaders[slug];
    if (!loader) {
      setMissing(true);
      return;
    }
    loader().then((raw) => {
      if (alive) setDoc(renderDoc(raw, slug));
    });
    return () => {
      alive = false;
    };
  }, [slug]);

  // Close the mobile nav on navigation, then scroll to anchor or top.
  useEffect(() => {
    setMenuOpen(false);
  }, [slug]);

  useEffect(() => {
    if (!doc) return;
    const id = location.hash.replace(/^#/, '');
    const el = id ? document.getElementById(id) : null;
    if (el) el.scrollIntoView();
    else window.scrollTo(0, 0);
  }, [doc, location.hash]);

  // Render any mermaid diagrams once the markdown is in the DOM.
  useEffect(() => {
    if (doc?.hasMermaid && articleRef.current) {
      void renderMermaid(articleRef.current);
    }
  }, [doc]);

  // Route in-app links embedded in the rendered markdown through the router.
  const onContentClick = (e: React.MouseEvent) => {
    const a = (e.target as HTMLElement).closest('a');
    if (!a) return;
    const href = a.getAttribute('href') ?? '';
    if (href.startsWith('/')) {
      e.preventDefault();
      nav(href);
    }
  };

  const idx = ORDER.indexOf(slug);
  const prev = idx > 0 ? ORDER[idx - 1] : null;
  const next = idx >= 0 && idx < ORDER.length - 1 ? ORDER[idx + 1] : null;

  return (
    <div className="docs">
      <DocsHeader onMenu={() => setMenuOpen((v) => !v)} menuOpen={menuOpen} />
      <div className="docs-shell">
        <Sidebar current={slug} open={menuOpen} onClose={() => setMenuOpen(false)} />
        <main className="docs-main">
          {missing ? (
            <article className="doc">
              <h1>Page not found</h1>
              <p>
                That documentation page doesn't exist. Head back to the{' '}
                <Link to="/docs">introduction</Link>.
              </p>
            </article>
          ) : !doc ? (
            <article className="doc">
              <div className="doc-loading">Loading…</div>
            </article>
          ) : (
            <>
              <article
                ref={articleRef}
                className="doc"
                onClick={onContentClick}
                dangerouslySetInnerHTML={{ __html: doc.html }}
              />
              <nav className="doc-pager">
                {prev !== null ? (
                  <Link className="pg prev" to={`/docs${prev ? '/' + prev : ''}`}>
                    <span>← Previous</span>
                    <b>{titleFor(prev)}</b>
                  </Link>
                ) : (
                  <span />
                )}
                {next !== null ? (
                  <Link className="pg next" to={`/docs${next ? '/' + next : ''}`}>
                    <span>Next →</span>
                    <b>{titleFor(next)}</b>
                  </Link>
                ) : (
                  <span />
                )}
              </nav>
            </>
          )}
        </main>
        <Toc items={doc?.toc ?? []} />
      </div>
    </div>
  );
}

function DocsHeader({ onMenu, menuOpen }: { onMenu: () => void; menuOpen: boolean }) {
  return (
    <header className="docs-top">
      <div className="docs-top-in">
        <button
          className="menu-btn"
          onClick={onMenu}
          aria-label="Toggle navigation"
          aria-expanded={menuOpen}
        >
          <span />
          <span />
          <span />
        </button>
        <Link className="brand" to="/" aria-label="Spindle home">
          <img src="/spindle.svg" width="25" height="25" alt="" aria-hidden="true" />
          Spindle <small>Docs</small>
        </Link>
        <nav className="docs-top-nav">
          <Link to="/">Home</Link>
          <a href={`${REPO}`} target="_blank" rel="noopener noreferrer">
            GitHub ↗
          </a>
        </nav>
      </div>
    </header>
  );
}

function Sidebar({
  current,
  open,
  onClose,
}: {
  current: string;
  open: boolean;
  onClose: () => void;
}) {
  return (
    <>
      {open && <div className="sidebar-scrim" onClick={onClose} />}
      <aside className={`docs-side${open ? ' open' : ''}`}>
        <nav>
          {NAV.map((group) => (
            <div className="side-group" key={group.label}>
              <div className="side-label">{group.label}</div>
              {group.items.map((item) => {
                const to = `/docs${item.slug ? '/' + item.slug : ''}`;
                const active = item.slug === current;
                return (
                  <Link key={to} to={to} className={`side-link${active ? ' active' : ''}`}>
                    {item.title}
                  </Link>
                );
              })}
            </div>
          ))}
        </nav>
      </aside>
    </>
  );
}

function Toc({ items }: { items: TocItem[] }) {
  if (items.length < 2) return <aside className="docs-toc" />;
  return (
    <aside className="docs-toc">
      <div className="toc-inner">
        <div className="toc-label">On this page</div>
        {items.map((it) => (
          <a key={it.id} href={`#${it.id}`} className={`toc-link lv${it.level}`}>
            {it.text}
          </a>
        ))}
      </div>
    </aside>
  );
}

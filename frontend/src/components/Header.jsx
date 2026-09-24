import { Link, NavLink } from 'react-router-dom';

export default function Header() {
  return (
    <header className="topbar">
      <div className="topbar-left">
        <Link to="/" className="brand">
          <svg width="22" height="22" viewBox="0 0 22 22" aria-hidden="true">
            <rect x="1" y="1" width="12" height="12" fill="#000" />
            <rect x="9" y="9" width="12" height="12" fill="none" stroke="var(--accent)" strokeWidth="2" />
          </svg>
          <span>LedgerLink</span>
        </Link>
        <nav className="nav">
          <NavLink to="/" end>Clients</NavLink>
        </nav>
      </div>
    </header>
  );
}

const base = { fill: 'none', stroke: 'currentColor', strokeWidth: 2, 'aria-hidden': true };

export const Plus = (p) => (<svg width="16" height="16" viewBox="0 0 24 24" {...base} {...p}><path d="M12 5v14M5 12h14" /></svg>);
export const Search = (p) => (<svg width="16" height="16" viewBox="0 0 24 24" {...base} {...p}><circle cx="11" cy="11" r="7" /><path d="M20 20l-3.5-3.5" /></svg>);
export const Trash = (p) => (<svg width="16" height="16" viewBox="0 0 24 24" {...base} {...p}><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3" /></svg>);
export const Check = (p) => (<svg width="13" height="13" viewBox="0 0 24 24" {...base} strokeWidth="3" {...p}><path d="M5 12l5 5 9-10" /></svg>);
export const Cross = (p) => (<svg width="16" height="16" viewBox="0 0 24 24" {...base} strokeWidth="2.5" {...p}><path d="M6 6l12 12M18 6L6 18" /></svg>);
export const Download = (p) => (<svg width="16" height="16" viewBox="0 0 24 24" {...base} {...p}><path d="M12 4v11M7 10l5 5 5-5M5 20h14" /></svg>);
export const Upload = (p) => (<svg width="16" height="16" viewBox="0 0 24 24" {...base} {...p}><path d="M12 20V9M7 14l5-5 5 5M5 4h14" /></svg>);
export const FileIcon = (p) => (<svg width="20" height="20" viewBox="0 0 24 24" {...base} strokeWidth="1.8" {...p}><path d="M14 3H6v18h12V7z" /><path d="M14 3v4h4" /></svg>);
export const WorkbookIcon = (p) => (<svg width="30" height="30" viewBox="0 0 24 24" {...base} strokeWidth="1.5" {...p}><rect x="3" y="6" width="14" height="15" rx="1.5" /><path d="M7 3h13v15" /><path d="M6 11h8M6 14h8M6 17h5" /></svg>);
export const GridIcon = (p) => (<svg width="30" height="30" viewBox="0 0 24 24" {...base} strokeWidth="1.5" {...p}><rect x="3" y="3" width="18" height="18" rx="1.5" /><path d="M3 9h18M3 15h18M12 3v18" /></svg>);
export const LinkIcon = (p) => (<svg width="14" height="14" viewBox="0 0 24 24" {...base} {...p}><path d="M10 14a4 4 0 0 0 5.66 0l3-3a4 4 0 0 0-5.66-5.66l-1 1" /><path d="M14 10a4 4 0 0 0-5.66 0l-3 3a4 4 0 0 0 5.66 5.66l1-1" /></svg>);

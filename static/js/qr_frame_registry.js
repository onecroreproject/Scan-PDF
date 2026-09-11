window.QR_FRAME_REGISTRY = {
    none: { label: 'None', kind: 'none' },
    simple_circle: { label: 'Circle', kind: 'border', radius: 'circle' },
    rounded_square: { label: 'Rounded Square', kind: 'border', radius: 'rounded' },
    simple_square: { label: 'Square', kind: 'border', radius: 'square' },
    bottom_bar: { label: 'Bottom Bar', kind: 'bar', side: 'bottom' },
    top_bar: { label: 'Top Bar', kind: 'bar', side: 'top' },
    top_bottom_bar: { label: 'Top + Bottom', kind: 'bars' },
    label_bottom: { label: 'Label Bottom', kind: 'label', side: 'bottom' },
    label_top: { label: 'Label Top', kind: 'label', side: 'top' },
    banner_top: { label: 'Banner Top', kind: 'banner', side: 'top' },
    banner_bottom: { label: 'Banner Bottom', kind: 'banner', side: 'bottom' },
    speech_top: { label: 'Speech Top', kind: 'speech', side: 'top' },
    speech_bottom: { label: 'Speech Bottom', kind: 'speech', side: 'bottom' },
    corners: { label: 'Corners', kind: 'corners' },
    rounded_corners: { label: 'Rounded Corners', kind: 'rounded_corners' },
    scan_me_top: { label: 'Scan Me Top', kind: 'banner', side: 'top' },
    scan_me_bottom: { label: 'Scan Me Bottom', kind: 'banner', side: 'bottom' },
    ribbon_top: { label: 'Ribbon Top', kind: 'ribbon', side: 'top' },
    ribbon_bottom: { label: 'Ribbon Bottom', kind: 'ribbon', side: 'bottom' },
    ticket: { label: 'Ticket', kind: 'ticket' },
    receipt: { label: 'Receipt', kind: 'receipt' },
    device_phone: { label: 'Phone', kind: 'device', device: 'phone' },
    device_tablet: { label: 'Tablet', kind: 'device', device: 'tablet' },
    badge: { label: 'Badge', kind: 'badge' },
    card: { label: 'Card', kind: 'card' },
    poster: { label: 'Poster', kind: 'poster' }
};

window.qrFrameIcon = function(type) {
    const f = QR_FRAME_REGISTRY[type] || QR_FRAME_REGISTRY.none;
    const stroke = '#6366f1';
    if (f.kind === 'none') return '<path d="M7 7 L17 17 M17 7 L7 17" stroke="#1e293b" stroke-width="2.4" stroke-linecap="round"/>';
    if (f.kind === 'border') return `<rect x="4" y="3" width="16" height="18" rx="${f.radius === 'circle' ? 8 : f.radius === 'rounded' ? 3 : 0}" fill="none" stroke="${stroke}" stroke-width="1.5"/><rect x="7" y="6" width="10" height="9" fill="#c7d2fe"/>`;
    if (f.kind === 'bar' || f.kind === 'banner' || f.kind === 'ribbon') return `<rect x="4" y="4" width="16" height="16" fill="none" stroke="#94a3b8"/><rect x="3" y="${f.side === 'top' ? 3 : 16}" width="18" height="5" rx="1" fill="${stroke}"/><rect x="7" y="8" width="10" height="6" fill="#c7d2fe"/>`;
    if (f.kind === 'bars') return '<rect x="3" y="3" width="18" height="18" fill="none" stroke="#94a3b8"/><rect x="3" y="3" width="18" height="4" fill="#6366f1"/><rect x="3" y="17" width="18" height="4" fill="#6366f1"/>';
    if (f.kind === 'speech') return `<path d="M4 4 H20 V16 H13 L10 20 V16 H4 Z" fill="none" stroke="${stroke}" stroke-width="1.5"/><rect x="7" y="7" width="10" height="6" fill="#c7d2fe"/>`;
    if (f.kind === 'corners' || f.kind === 'rounded_corners') return '<path d="M4 9 V4 H9 M15 4 H20 V9 M20 15 V20 H15 M9 20 H4 V15" fill="none" stroke="#6366f1" stroke-width="2" stroke-linecap="round"/><rect x="7" y="7" width="10" height="10" fill="#c7d2fe"/>';
    if (f.kind === 'ticket' || f.kind === 'receipt') return '<path d="M5 3 H19 V21 H5 Z" fill="none" stroke="#6366f1"/><rect x="7" y="6" width="10" height="8" fill="#c7d2fe"/><path d="M7 17 H17" stroke="#6366f1"/>';
    if (f.kind === 'device') return `<rect x="${f.device === 'phone' ? 7 : 4}" y="3" width="${f.device === 'phone' ? 10 : 16}" height="18" rx="2" fill="none" stroke="#6366f1"/><rect x="7" y="6" width="10" height="9" fill="#c7d2fe"/>`;
    return '<rect x="4" y="3" width="16" height="18" rx="2" fill="none" stroke="#6366f1"/><rect x="7" y="6" width="10" height="9" fill="#c7d2fe"/><path d="M7 18 H17" stroke="#6366f1" stroke-width="2"/>';
};

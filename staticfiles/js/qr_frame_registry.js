window.QR_FRAME_REGISTRY = {
    none: { label: 'None', kind: 'none' },
    simple_circle: { label: 'Circle', kind: 'border', radius: 'circle' },
    rounded_square: { label: 'Rounded Square', kind: 'border', radius: 'rounded' },
    simple_square: { label: 'Square', kind: 'border', radius: 'square' },
    bottom_bar: { label: 'Bottom Bar', kind: 'bar', side: 'bottom', text: true },
    top_bar: { label: 'Top Bar', kind: 'bar', side: 'top', text: true },
    top_bottom_bar: { label: 'Top + Bottom', kind: 'bars', text: true },
    label_bottom: { label: 'Label Bottom', kind: 'label', side: 'bottom', text: true },
    label_top: { label: 'Label Top', kind: 'label', side: 'top', text: true },
    banner_top: { label: 'Banner Top', kind: 'banner', side: 'top', text: true },
    banner_bottom: { label: 'Banner Bottom', kind: 'banner', side: 'bottom', text: true },
    speech_top: { label: 'Speech Top', kind: 'speech', side: 'top', text: true },
    speech_bottom: { label: 'Speech Bottom', kind: 'speech', side: 'bottom', text: true },
    corners: { label: 'Corners', kind: 'corners' },
    rounded_corners: { label: 'Rounded Corners', kind: 'rounded_corners' },
    scan_me_top: { label: 'Scan Me Top', kind: 'banner', side: 'top', text: true },
    scan_me_bottom: { label: 'Scan Me Bottom', kind: 'banner', side: 'bottom', text: true },
    ribbon_top: { label: 'Ribbon Top', kind: 'ribbon', side: 'top', text: true },
    ribbon_bottom: { label: 'Ribbon Bottom', kind: 'ribbon', side: 'bottom', text: true },
    ticket: { label: 'Ticket', kind: 'ticket', text: true },
    receipt: { label: 'Receipt', kind: 'receipt', text: true },
    device_phone: { label: 'Phone', kind: 'device', device: 'phone' },
    device_tablet: { label: 'Tablet', kind: 'device', device: 'tablet' },
    badge: { label: 'Badge', kind: 'badge', text: true },
    card: { label: 'Card', kind: 'card', text: true },
    poster: { label: 'Poster', kind: 'poster', text: true }
};

window.renderQRFrameSVG = function(config) {
    const { type, text = 'SCAN ME', color = '#6366f1', textColor = '#ffffff' } = config;
    const frame_type = type;
    
    // Base dimensions
    const cell = 12;
    const modules = 29; // Dummy QR size
    const img_px = modules * cell; // 348
    const quiet_zone = 24;
    const base_qr_w = img_px + quiet_zone * 2; // 396
    
    let top_extra = 0, bottom_extra = 0, left_extra = 0, right_extra = 0;
    let frame_padding = 16;
    
    let frame_is_active = frame_type !== 'none';
    if (frame_is_active) {
        if (frame_type === 'top_bar') top_extra = 60;
        else if (frame_type === 'bottom_bar') bottom_extra = 60;
        else if (frame_type === 'top_bottom_bar') { top_extra = 55; bottom_extra = 55; }
        else if (frame_type === 'label_top') top_extra = 70;
        else if (frame_type === 'label_bottom') bottom_extra = 70;
        else if (frame_type === 'speech_top') top_extra = 85;
        else if (frame_type === 'speech_bottom') bottom_extra = 85;
        else if (frame_type === 'ribbon_top') { top_extra = 80; left_extra = 12; right_extra = 12; }
        else if (frame_type === 'ribbon_bottom') { bottom_extra = 80; left_extra = 12; right_extra = 12; }
        else if (frame_type === 'scan_me_top' || frame_type === 'banner_top') top_extra = 70;
        else if (frame_type === 'scan_me_bottom' || frame_type === 'banner_bottom') bottom_extra = 70;
        else if (frame_type === 'ticket' || frame_type === 'receipt') { left_extra = 18; right_extra = 18; bottom_extra = 70; }
        else if (frame_type === 'device_phone') { top_extra = 60; bottom_extra = 80; left_extra = 16; right_extra = 16; }
        else if (frame_type === 'device_tablet') { top_extra = 45; bottom_extra = 60; left_extra = 30; right_extra = 30; }
        else if (frame_type === 'badge') { bottom_extra = 70; left_extra = 16; right_extra = 16; }
        else if (frame_type === 'card') bottom_extra = 70;
        else if (frame_type === 'poster') { bottom_extra = 90; top_extra = 30; }
        else if (frame_type === 'simple_circle') frame_padding = Math.floor(base_qr_w * 0.15);
    } else {
        frame_padding = 0;
    }
    
    const frame_w = base_qr_w + (frame_padding * 2);
    const output_w = frame_w + left_extra + right_extra;
    const output_h = base_qr_w + (frame_padding * 2) + top_extra + bottom_extra;
    const qr_offset_x = left_extra + frame_padding + quiet_zone;
    const qr_offset_y = top_extra + frame_padding + quiet_zone;
    
    const fx = left_extra;
    const fy = top_extra;
    const fw = frame_w;
    const fh = base_qr_w + frame_padding * 2;
    
    let svg = `<svg viewBox="0 0 ${output_w} ${output_h}" xmlns="http://www.w3.org/2000/svg" width="100%" height="100%">`;
    
    // Draw dummy QR
    const qrFill = "#1e293b";
    svg += `<defs>
        <pattern id="qr-pattern-${type}" width="48" height="48" patternUnits="userSpaceOnUse">
            <rect x="0" y="0" width="24" height="24" fill="${qrFill}"/>
            <rect x="24" y="24" width="24" height="24" fill="${qrFill}"/>
            <rect x="0" y="24" width="24" height="24" fill="${qrFill}" opacity="0.5"/>
        </pattern>
    </defs>`;
    
    let qr_elements = `
        <rect x="0" y="0" width="${img_px}" height="${img_px}" fill="none" rx="8"/>
        <!-- Eyes -->
        <rect x="0" y="0" width="84" height="84" fill="none" stroke="${qrFill}" stroke-width="12" rx="12"/>
        <rect x="24" y="24" width="36" height="36" fill="${qrFill}" rx="6"/>
        
        <rect x="${img_px-84}" y="0" width="84" height="84" fill="none" stroke="${qrFill}" stroke-width="12" rx="12"/>
        <rect x="${img_px-60}" y="24" width="36" height="36" fill="${qrFill}" rx="6"/>
        
        <rect x="0" y="${img_px-84}" width="84" height="84" fill="none" stroke="${qrFill}" stroke-width="12" rx="12"/>
        <rect x="24" y="${img_px-60}" width="36" height="36" fill="${qrFill}" rx="6"/>
        
        <rect x="108" y="108" width="${img_px-108}" height="${img_px-108}" fill="url(#qr-pattern-${type})"/>
    `;
    
    if (frame_is_active) {
        let radius = (frame_type === 'rounded_square' || frame_type === 'rounded_corners' || frame_type === 'card') ? 24 : 0;
        
        if (frame_type === 'simple_circle') {
            svg += `<circle cx="${output_w/2}" cy="${output_h/2}" r="${frame_w/2}" fill="none" stroke="${color}" stroke-width="12"/>`;
        } else if (frame_type === 'simple_square' || frame_type === 'rounded_square') {
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh}" rx="${radius}" fill="none" stroke="${color}" stroke-width="12"/>`;
        } else if (frame_type === 'card') {
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh + bottom_extra}" rx="28" fill="${color}"/>`;
            svg += `<rect x="${fx+12}" y="${fy+12}" width="${fw-24}" height="${fh-24}" rx="20" fill="#ffffff"/>`;
        } else if (frame_type === 'poster') {
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh + bottom_extra + top_extra}" fill="none" stroke="${color}" stroke-width="12"/>`;
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${top_extra + 10}" fill="${color}"/>`;
        } else if (frame_type === 'badge') {
            svg += `<path d="M${fx} ${fy} H${fx+fw} V${fy+fh} L${output_w/2} ${output_h} L${fx} ${fy+fh} Z" fill="none" stroke="${color}" stroke-width="12" stroke-linejoin="round"/>`;
            svg += `<path d="M${fx+12} ${fy+12} H${fx+fw-12} V${fy+fh-8} L${output_w/2} ${output_h-18} L${fx+12} ${fy+fh-8} Z" fill="${color}" opacity="0.1"/>`;
        } else if (frame_type === 'corners' || frame_type === 'rounded_corners') {
            let c = 60;
            let p1 = `M${fx} ${fy+c} V${fy} H${fx+c}`;
            let p2 = `M${fx+fw-c} ${fy} H${fx+fw} V${fy+c}`;
            let p3 = `M${fx+fw} ${fy+fh-c} V${fy+fh} H${fx+fw-c}`;
            let p4 = `M${fx+c} ${fy+fh} H${fx} V${fy+fh-c}`;
            svg += `<path d="${p1} ${p2} ${p3} ${p4}" fill="none" stroke="${color}" stroke-width="12" stroke-linecap="round"/>`;
        } else if (frame_type === 'ticket' || frame_type === 'receipt') {
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh + bottom_extra}" fill="none" stroke="${color}" stroke-width="8" rx="12"/>`;
            if (frame_type === 'ticket') {
                svg += `<circle cx="${fx}" cy="${output_h/2}" r="16" fill="#ffffff" stroke="${color}" stroke-width="8"/>`;
                svg += `<circle cx="${fx+fw}" cy="${output_h/2}" r="16" fill="#ffffff" stroke="${color}" stroke-width="8"/>`;
            }
            svg += `<line x1="${fx+24}" y1="${fy+fh}" x2="${fx+fw-24}" y2="${fy+fh}" stroke="${color}" stroke-width="6" stroke-dasharray="12 12"/>`;
        } else if (frame_type === 'device_phone') {
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh + bottom_extra + top_extra}" rx="48" fill="none" stroke="${color}" stroke-width="12"/>`;
            svg += `<rect x="${output_w/2-40}" y="${fy+24}" width="80" height="8" rx="4" fill="${color}"/>`;
        } else if (frame_type === 'device_tablet') {
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh + bottom_extra + top_extra}" rx="24" fill="none" stroke="${color}" stroke-width="12"/>`;
            svg += `<circle cx="${output_w/2}" cy="${fy+24}" r="8" fill="${color}"/>`;
        } else if (frame_type === 'speech_top' || frame_type === 'speech_bottom') {
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh}" rx="20" fill="none" stroke="${color}" stroke-width="8"/>`;
            let lbl_w = Math.min(fw * 0.85, Math.max(160, text.length * 12 + 60));
            let lx = output_w/2 - lbl_w/2;
            if (frame_type === 'speech_top') {
                let lh = top_extra - 25, ly = fy - top_extra + 10;
                svg += `<rect x="${lx}" y="${ly}" width="${lbl_w}" height="${lh}" rx="${lh/2}" fill="${color}"/>`;
                svg += `<polygon points="${output_w/2-16},${ly+lh-2} ${output_w/2+16},${ly+lh-2} ${output_w/2},${ly+lh+18}" fill="${color}"/>`;
            } else {
                let lh = bottom_extra - 25, ly = fy + fh + 15;
                svg += `<rect x="${lx}" y="${ly}" width="${lbl_w}" height="${lh}" rx="${lh/2}" fill="${color}"/>`;
                svg += `<polygon points="${output_w/2-16},${ly+2} ${output_w/2+16},${ly+2} ${output_w/2},${ly-18}" fill="${color}"/>`;
            }
        } else if (frame_type === 'ribbon_top' || frame_type === 'ribbon_bottom') {
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh}" fill="none" stroke="${color}" stroke-width="8"/>`;
            let rw = fw * 0.9, rh = 56, fold = 18;
            let ry = frame_type === 'ribbon_top' ? fy - rh/2 : fy + fh - rh/2;
            let rx = output_w/2 - rw/2;
            svg += `<polygon points="${rx-left_extra},${ry+fold} ${rx},${ry} ${rx},${ry+rh} ${rx-left_extra},${ry+rh-fold}" fill="${color}" opacity="0.8"/>`;
            svg += `<polygon points="${rx+rw+right_extra},${ry+fold} ${rx+rw},${ry} ${rx+rw},${ry+rh} ${rx+rw+right_extra},${ry+rh-fold}" fill="${color}" opacity="0.8"/>`;
            svg += `<rect x="${rx}" y="${ry}" width="${rw}" height="${rh}" fill="${color}"/>`;
        } else if (frame_type === 'label_top' || frame_type === 'label_bottom') {
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh}" rx="20" fill="none" stroke="${color}" stroke-width="8"/>`;
            let lbl_w = Math.min(fw * 0.85, Math.max(160, text.length * 12 + 60));
            let lx = output_w/2 - lbl_w/2, lh = 56;
            if (frame_type === 'label_top') {
                svg += `<rect x="${lx}" y="${fy - lh - 12}" width="${lbl_w}" height="${lh}" rx="${lh/2}" fill="${color}"/>`;
            } else {
                svg += `<rect x="${lx}" y="${fy + fh + 12}" width="${lbl_w}" height="${lh}" rx="${lh/2}" fill="${color}"/>`;
            }
        } else if (['top_bar', 'bottom_bar', 'top_bottom_bar', 'scan_me_top', 'scan_me_bottom', 'banner_top', 'banner_bottom'].includes(frame_type)) {
            svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh}" rx="12" fill="none" stroke="${color}" stroke-width="8"/>`;
            if (top_extra > 0) {
                svg += `<path d="M${fx} ${fy+top_extra} V${fy+12} A12 12 0 0 1 ${fx+12} ${fy} H${fx+fw-12} A12 12 0 0 1 ${fx+fw} ${fy} V${fy+top_extra} Z" fill="${color}"/>`;
            }
            if (bottom_extra > 0) {
                svg += `<path d="M${fx} ${fy+fh-bottom_extra} V${fy+fh-12} A12 12 0 0 0 ${fx+12} ${fy+fh} H${fx+fw-12} A12 12 0 0 0 ${fx+fw} ${fy+fh-12} V${fy+fh-bottom_extra} Z" fill="${color}"/>`;
            }
        }

        let ty = 0;
        let ts = 28;
        if (['top_bar', 'scan_me_top', 'banner_top', 'top_bottom_bar', 'poster'].includes(frame_type)) ty = fy + top_extra/2;
        else if (['bottom_bar', 'scan_me_bottom', 'banner_bottom', 'top_bottom_bar'].includes(frame_type)) ty = fy + fh - bottom_extra/2;
        else if (frame_type === 'label_top') ty = fy - 56/2 - 12;
        else if (frame_type === 'label_bottom') ty = fy + fh + 56/2 + 12;
        else if (frame_type === 'speech_top') ty = fy - top_extra + 10 + (top_extra-25)/2;
        else if (frame_type === 'speech_bottom') ty = fy + fh + 15 + (bottom_extra-25)/2;
        else if (frame_type === 'ribbon_top') ty = fy;
        else if (frame_type === 'ribbon_bottom') ty = fy + fh;
        else if (['ticket', 'receipt', 'card'].includes(frame_type)) { ty = fy + fh + bottom_extra/2; }
        else if (frame_type === 'badge') ty = fy + fh + bottom_extra/2 - 12;
        
        let finalTextColor = ['card'].includes(frame_type) ? '#ffffff' : textColor;
        let fillStyle = ['badge', 'receipt', 'ticket'].includes(frame_type) ? color : finalTextColor;
        
        if (ty > 0) {
            svg += `<text x="${output_w/2}" y="${ty}" font-family="system-ui, sans-serif" font-size="${ts}" font-weight="bold" fill="${fillStyle}" text-anchor="middle" dominant-baseline="central">${text.substring(0, 15)}</text>`;
        }
    }
    
    svg += `<g transform="translate(${qr_offset_x},${qr_offset_y})">${qr_elements}</g>`;
    svg += `</svg>`;
    
    return svg;
};

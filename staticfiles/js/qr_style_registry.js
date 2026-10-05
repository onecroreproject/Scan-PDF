/* Shared QR style metadata for the create and edit studios. */
window.QR_STYLE_REGISTRY = {
    body: [
        ['square', 'Square'], ['rounded_square', 'Rounded Square'], ['circle', 'Circle'],
        ['diamond', 'Diamond'], ['small_circle', 'Small Circle'], ['small_square', 'Small Square'],
        ['horizontal', 'Horizontal'], ['vertical', 'Vertical'], ['star', 'Star'], ['plus', 'Plus'],
        ['rounded', 'Rounded'], ['connected', 'Connected'], ['hexagon', 'Hexagon'], ['octagon', 'Octagon'],
        ['cross', 'Cross'], ['flower', 'Flower'], ['clover', 'Clover'], ['sparkle', 'Sparkle'],
        ['soft_diamond', 'Soft Diamond'], ['burst', 'Burst'], ['droplet', 'Droplet'], ['leaf', 'Leaf'],
        ['angled_square', 'Angled Square'], ['tilted_square', 'Tilted Square'], ['four_dots', 'Four Dots'],
        ['pixel_plus', 'Pixel Plus'], ['soft_cross', 'Soft Cross'], ['corner_round', 'Corner Round'],
        ['blob', 'Blob']
    ],
    outer_eye: [
        ['square', 'Square'], ['circle', 'Circle'], ['rounded_square', 'Rounded Square'],
        ['diamond', 'Diamond'], ['soft_square', 'Soft Square'], ['leaf', 'Leaf'], ['hexagon', 'Hexagon'],
        ['octagon', 'Octagon'], ['star', 'Star'], ['cut_corner', 'Cut Corner'], ['blob', 'Blob'],
        ['squircle', 'Squircle']
    ],
    inner_eye: [
        ['square', 'Square'], ['circle', 'Circle'], ['rounded_square', 'Rounded Square'],
        ['diamond', 'Diamond'], ['soft_square', 'Soft Square'], ['leaf', 'Leaf'], ['star', 'Star'],
        ['hexagon', 'Hexagon'], ['octagon', 'Octagon'], ['clover', 'Clover'], ['sparkle', 'Sparkle'],
        ['blob', 'Blob']
    ]
};

window.qrStyleIcon = function(style, color, outline) {
    const fill = outline ? 'none' : color;
    const stroke = outline ? color : 'none';
    const sw = outline ? 1.8 : 0;
    const common = `fill="${fill}" stroke="${stroke}" stroke-width="${sw}"`;
    const shape = {
        square: `<rect x="5" y="5" width="14" height="14" ${common} rx="1"/>`,
        rounded_square: `<rect x="5" y="5" width="14" height="14" ${common} rx="4"/>`,
        circle: `<circle cx="12" cy="12" r="7" ${common}/>`,
        diamond: `<path d="M12 4 L20 12 L12 20 L4 12 Z" ${common}/>`,
        small_circle: `<circle cx="12" cy="12" r="4" ${common}/>`,
        small_square: `<rect x="8" y="8" width="8" height="8" ${common}/>`,
        horizontal: `<rect x="4" y="10" width="16" height="4" ${common} rx="1"/>`,
        vertical: `<rect x="10" y="4" width="4" height="16" ${common} rx="1"/>`,
        star: `<path d="M12 3 L14.2 9.2 L21 9.3 L15.6 13.2 L17.5 20 L12 16 L6.5 20 L8.4 13.2 L3 9.3 L9.8 9.2 Z" ${common}/>`,
        plus: `<path d="M9 3 H15 V9 H21 V15 H15 V21 H9 V15 H3 V9 H9 Z" ${common}/>`,
        rounded: `<rect x="4" y="4" width="16" height="16" ${common} rx="6"/>`,
        connected: `<path d="M5 5 H11 V11 H13 V5 H19 V11 H13 V13 H19 V19 H13 V13 H11 V19 H5 V13 H11 V11 H5 Z" ${common}/>`,
        hexagon: `<path d="M7 4 H17 L21 12 L17 20 H7 L3 12 Z" ${common}/>`,
        octagon: `<path d="M8 3 H16 L21 8 V16 L16 21 H8 L3 16 V8 Z" ${common}/>`,
        cross: `<path d="M9 3 H15 V9 H21 V15 H15 V21 H9 V15 H3 V9 H9 Z" ${common}/>`,
        flower: `<path d="M12 12 C7 12 5 10 5 7 C5 4 8 3 10 5 C12 0 16 1 16 5 C20 3 22 6 20 9 C24 12 21 16 17 15 C18 20 14 22 11 18 C7 21 4 17 6 14 C2 13 3 9 6 8 C7 10 9 12 12 12 Z" ${common}/>`,
        clover: `<path d="M12 12 C7 12 5 10 5 7 C5 4 8 3 10 6 C10 2 14 2 14 6 C17 3 20 5 19 8 C23 8 23 12 19 13 C22 16 19 19 16 18 C16 22 12 22 12 18 C9 21 6 19 7 16 C3 16 3 12 7 12 Z" ${common}/>`,
        sparkle: `<path d="M12 2 L14 10 L22 12 L14 14 L12 22 L10 14 L2 12 L10 10 Z" ${common}/>`,
        soft_diamond: `<path d="M12 3 Q14 7 21 12 Q14 17 12 21 Q10 17 3 12 Q10 7 12 3 Z" ${common}/>`,
        burst: `<path d="M12 2 L14 7 L18 4 L17 9 L22 9 L18 12 L22 15 L17 15 L18 20 L14 17 L12 22 L10 17 L6 20 L7 15 L2 15 L6 12 L2 9 L7 9 L6 4 L10 7 Z" ${common}/>`,
        droplet: `<path d="M12 3 C12 3 19 10 19 15 A7 7 0 1 1 5 15 C5 10 12 3 12 3 Z" ${common}/>`,
        leaf: `<path d="M20 4 C9 4 4 9 4 16 C4 20 8 20 12 18 C17 15 20 10 20 4 Z" ${common}/>`,
        angled_square: `<path d="M7 4 H20 V17 L17 20 H4 V7 Z" ${common}/>`,
        tilted_square: `<path d="M6 5 L19 3 L18 18 L5 20 Z" ${common}/>`,
        four_dots: `<circle cx="8" cy="8" r="3" ${common}/><circle cx="16" cy="8" r="3" ${common}/><circle cx="8" cy="16" r="3" ${common}/><circle cx="16" cy="16" r="3" ${common}/>`,
        pixel_plus: `<path d="M8 3 H16 V8 H21 V16 H16 V21 H8 V16 H3 V8 H8 Z" ${common}/>`,
        soft_cross: `<path d="M8 4 Q12 7 16 4 Q17 8 20 12 Q17 16 20 20 Q16 17 12 20 Q8 17 4 20 Q7 16 4 12 Q7 8 4 4 Q8 7 8 4 Z" ${common}/>`,
        corner_round: `<path d="M9 4 H20 V15 Q20 20 15 20 H4 V9 Q4 4 9 4 Z" ${common} rx="2"/>`,
        blob: `<path d="M12 3 C17 3 21 7 20 12 C21 17 17 21 12 20 C7 21 3 17 4 12 C3 7 7 3 12 3 Z" ${common}/>`,
        soft_square: `<rect x="4" y="4" width="16" height="16" ${common} rx="3"/>`,
        cut_corner: `<path d="M7 4 H20 V17 L17 20 H4 V7 Z" ${common}/>`,
        squircle: `<rect x="4" y="4" width="16" height="16" ${common} rx="7"/>`
    };
    return shape[style] || shape.square;
};

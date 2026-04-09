(() => {
    const categories = [
        { key: 'area', title: 'Area Converter', subtitle: 'Convert area units using square meter as the base unit.', icon: '▦', type: 'standard' },
        { key: 'currency', title: 'Currency Converter', subtitle: 'Live rates (if configured) with static fallback and graceful failure handling.', icon: '$', type: 'currency' },
        { key: 'length', title: 'Length / Distance Converter', subtitle: 'Convert distance units with meter as the base unit.', icon: '↔', type: 'standard' },
        { key: 'speed', title: 'Speed Converter', subtitle: 'Convert speed units with meter per second as the base unit.', icon: '⚡', type: 'standard' },
        { key: 'temperature', title: 'Temperature Converter', subtitle: 'Formula-based conversion for Celsius, Fahrenheit, and Kelvin.', icon: '°', type: 'temperature' },
        { key: 'time', title: 'Time Converter', subtitle: 'Time conversion using second as base. Month/year are approximate.', icon: '⏱', type: 'standard' },
        { key: 'timezone', title: 'Time Zone Converter', subtitle: 'Timezone-aware conversion using real IANA zones and DST-aware browser Intl.', icon: '🌐', type: 'timezone' },
        { key: 'volume', title: 'Volume / Capacity Converter', subtitle: 'Convert volume units with liter as the base unit.', icon: '🧪', type: 'standard' },
        { key: 'weight', title: 'Weight / Mass Converter', subtitle: 'Convert mass units with kilogram as the base unit.', icon: '⚖', type: 'standard' },
    ];

    const units = {
        area: {
            sqmm: { label: 'Square millimeter', factor: 0.000001 },
            sqcm: { label: 'Square centimeter', factor: 0.0001 },
            sqm: { label: 'Square meter', factor: 1 },
            sqkm: { label: 'Square kilometer', factor: 1000000 },
            sqin: { label: 'Square inch', factor: 0.00064516 },
            sqft: { label: 'Square foot', factor: 0.09290304 },
            sqyd: { label: 'Square yard', factor: 0.83612736 },
            acre: { label: 'Acre', factor: 4046.8564224 },
            hectare: { label: 'Hectare', factor: 10000 },
        },
        length: {
            mm: { label: 'Millimeter', factor: 0.001 },
            cm: { label: 'Centimeter', factor: 0.01 },
            m: { label: 'Meter', factor: 1 },
            km: { label: 'Kilometer', factor: 1000 },
            inch: { label: 'Inch', factor: 0.0254 },
            ft: { label: 'Foot', factor: 0.3048 },
            yd: { label: 'Yard', factor: 0.9144 },
            mile: { label: 'Mile', factor: 1609.344 },
            nmi: { label: 'Nautical mile', factor: 1852 },
        },
        speed: {
            ms: { label: 'Meter/second (m/s)', factor: 1 },
            kmh: { label: 'Kilometer/hour (km/h)', factor: 0.2777777778 },
            mph: { label: 'Mile/hour (mph)', factor: 0.44704 },
            knot: { label: 'Knot', factor: 0.5144444444 },
            fts: { label: 'Foot/second (ft/s)', factor: 0.3048 },
        },
        time: {
            ms: { label: 'Millisecond', factor: 0.001 },
            sec: { label: 'Second', factor: 1 },
            min: { label: 'Minute', factor: 60 },
            hr: { label: 'Hour', factor: 3600 },
            day: { label: 'Day', factor: 86400 },
            week: { label: 'Week', factor: 604800 },
            month: { label: 'Month (approx 30.44 days)', factor: 2629746 },
            year: { label: 'Year (approx 365.2425 days)', factor: 31556952 },
        },
        volume: {
            ml: { label: 'Milliliter', factor: 0.001 },
            l: { label: 'Liter', factor: 1 },
            m3: { label: 'Cubic meter', factor: 1000 },
            tsp: { label: 'Teaspoon', factor: 0.00492892159375 },
            tbsp: { label: 'Tablespoon', factor: 0.01478676478125 },
            cup: { label: 'Cup', factor: 0.2365882365 },
            pint: { label: 'Pint', factor: 0.473176473 },
            quart: { label: 'Quart', factor: 0.946352946 },
            gallon: { label: 'Gallon', factor: 3.785411784 },
        },
        weight: {
            mg: { label: 'Milligram', factor: 0.000001 },
            g: { label: 'Gram', factor: 0.001 },
            kg: { label: 'Kilogram', factor: 1 },
            ton: { label: 'Ton (metric)', factor: 1000 },
            oz: { label: 'Ounce', factor: 0.028349523125 },
            lb: { label: 'Pound', factor: 0.45359237 },
            stone: { label: 'Stone', factor: 6.35029318 },
        },
    };

    const temperatureUnits = {
        c: 'Celsius',
        f: 'Fahrenheit',
        k: 'Kelvin',
    };

    const currencyOrder = ['INR', 'USD', 'EUR', 'GBP', 'AED', 'JPY', 'SGD', 'AUD', 'CAD', 'CNY'];
    const fallbackCurrencyRatesUSD = {
        USD: 1,
        INR: 83.25,
        EUR: 0.92,
        GBP: 0.78,
        AED: 3.67,
        JPY: 151.4,
        SGD: 1.35,
        AUD: 1.52,
        CAD: 1.36,
        CNY: 7.24,
    };

    const timezones = [
        'Asia/Kolkata',
        'UTC',
        'America/New_York',
        'America/Los_Angeles',
        'Europe/London',
        'Europe/Paris',
        'Asia/Dubai',
        'Asia/Singapore',
        'Asia/Tokyo',
        'Australia/Sydney',
    ];

    const state = {
        current: 'length',
        currencyRates: { ...fallbackCurrencyRatesUSD },
        currencyLive: false,
        currencyLastUpdated: null,
    };

    const el = {
        sidebar: document.getElementById('uc-sidebar'),
        title: document.getElementById('uc-title'),
        subtitle: document.getElementById('uc-subtitle'),
        error: document.getElementById('uc-error'),
        badge: document.getElementById('currency-status'),
        standardGrid: document.getElementById('uc-standard-grid'),
        timezoneGrid: document.getElementById('uc-timezone-grid'),
        input: document.getElementById('uc-input'),
        from: document.getElementById('uc-from'),
        to: document.getElementById('uc-to'),
        tzDatetime: document.getElementById('tz-datetime'),
        tzFrom: document.getElementById('tz-from'),
        tzTo: document.getElementById('tz-to'),
        swap: document.getElementById('uc-swap'),
        reset: document.getElementById('uc-reset'),
        convert: document.getElementById('uc-convert'),
        result: document.getElementById('uc-result'),
        meta: document.getElementById('uc-meta'),
    };

    function formatNumber(num) {
        if (!Number.isFinite(num)) return '-';
        return Number(num).toLocaleString(undefined, { maximumFractionDigits: 8 });
    }

    function setError(message = '') {
        if (!message) {
            el.error.hidden = true;
            el.error.textContent = '';
            return;
        }
        el.error.hidden = false;
        el.error.textContent = message;
    }

    function setResult(value, meta = '') {
        el.result.textContent = value;
        el.meta.textContent = meta;
    }

    function renderSidebar() {
        el.sidebar.innerHTML = '';
        categories.forEach((cat) => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = `uc-cat-btn ${cat.key === state.current ? 'active' : ''}`;
            btn.textContent = `${cat.icon}  ${cat.title.replace(' Converter', '')}`;
            btn.addEventListener('click', () => switchCategory(cat.key));
            el.sidebar.appendChild(btn);
        });
    }

    function buildOptions(selectEl, optionsObj) {
        selectEl.innerHTML = '';
        Object.entries(optionsObj).forEach(([value, cfg]) => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = cfg.label;
            selectEl.appendChild(option);
        });
    }

    function buildCurrencyOptions(selectEl) {
        selectEl.innerHTML = '';
        currencyOrder.forEach((code) => {
            const option = document.createElement('option');
            option.value = code;
            option.textContent = code;
            selectEl.appendChild(option);
        });
    }

    function buildTimezoneOptions(selectEl) {
        selectEl.innerHTML = '';
        timezones.forEach((zone) => {
            const option = document.createElement('option');
            option.value = zone;
            option.textContent = zone;
            selectEl.appendChild(option);
        });
    }

    function categoryConfig(key) {
        return categories.find((cat) => cat.key === key);
    }

    function isNegativeAllowed(key) {
        return key === 'temperature';
    }

    async function switchCategory(key) {
        state.current = key;
        const cat = categoryConfig(key);
        el.title.textContent = cat.title;
        el.subtitle.textContent = cat.subtitle;
        setError('');
        setResult('-', '');

        const isTimezone = cat.type === 'timezone';
        el.standardGrid.hidden = isTimezone;
        el.timezoneGrid.hidden = !isTimezone;
        el.swap.disabled = false;

        if (cat.type === 'standard') {
            buildOptions(el.from, units[key]);
            buildOptions(el.to, units[key]);
            el.to.selectedIndex = Math.min(1, el.to.options.length - 1);
            el.input.value = '1';
        } else if (cat.type === 'temperature') {
            const map = Object.fromEntries(Object.entries(temperatureUnits).map(([k, v]) => [k, { label: v }]));
            buildOptions(el.from, map);
            buildOptions(el.to, map);
            el.to.selectedIndex = 1;
            el.input.value = '0';
        } else if (cat.type === 'currency') {
            buildCurrencyOptions(el.from);
            buildCurrencyOptions(el.to);
            el.from.value = 'USD';
            el.to.value = 'INR';
            el.input.value = '1';
            await loadCurrencyRates();
        } else if (cat.type === 'timezone') {
            buildTimezoneOptions(el.tzFrom);
            buildTimezoneOptions(el.tzTo);
            el.tzFrom.value = 'UTC';
            el.tzTo.value = 'Asia/Kolkata';
            el.tzDatetime.value = defaultDateTimeLocal();
        }

        updateCurrencyBadge();
        renderSidebar();
        convertCurrent();
    }

    function defaultDateTimeLocal() {
        const now = new Date();
        const pad = (x) => String(x).padStart(2, '0');
        return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}T${pad(now.getHours())}:${pad(now.getMinutes())}`;
    }

    function convertStandard() {
        const key = state.current;
        const amount = Number(el.input.value);
        if (!Number.isFinite(amount)) return setError('Please enter a valid numeric value.'), setResult('-', '');
        if (!isNegativeAllowed(key) && amount < 0) return setError('Negative values are not allowed for this converter.'), setResult('-', '');
        setError('');
        const fromCfg = units[key][el.from.value];
        const toCfg = units[key][el.to.value];
        const result = amount * (fromCfg.factor / toCfg.factor);
        setResult(`${formatNumber(result)} ${toCfg.label}`, `1 ${fromCfg.label} = ${formatNumber(fromCfg.factor / toCfg.factor)} ${toCfg.label}`);
    }

    function convertTemperature() {
        const amount = Number(el.input.value);
        if (!Number.isFinite(amount)) return setError('Please enter a valid numeric value.'), setResult('-', '');
        setError('');
        const from = el.from.value;
        const to = el.to.value;
        let celsius;
        if (from === 'c') celsius = amount;
        if (from === 'f') celsius = (amount - 32) * (5 / 9);
        if (from === 'k') celsius = amount - 273.15;
        let result = celsius;
        if (to === 'f') result = celsius * (9 / 5) + 32;
        if (to === 'k') result = celsius + 273.15;
        setResult(`${formatNumber(result)} ${temperatureUnits[to]}`, 'Formula-based conversion');
    }

    function convertCurrency() {
        const amount = Number(el.input.value);
        if (!Number.isFinite(amount)) return setError('Please enter a valid numeric value.'), setResult('-', '');
        if (amount < 0) return setError('Negative values are not allowed for currency conversion.'), setResult('-', '');
        const from = el.from.value;
        const to = el.to.value;
        const fromRate = state.currencyRates[from];
        const toRate = state.currencyRates[to];
        if (!fromRate || !toRate) return setError('Currency rates are unavailable right now.'), setResult('-', '');
        setError('');
        const result = amount * (toRate / fromRate);
        const source = state.currencyLive ? 'Live rates' : 'Static fallback rates';
        const stamp = state.currencyLive && state.currencyLastUpdated
            ? `Last updated: ${new Date(state.currencyLastUpdated * 1000).toLocaleString()}`
            : 'Last updated: fallback mode';
        setResult(`${formatNumber(result)} ${to}`, `${source} | ${stamp}`);
    }

    function getOffsetMs(timestamp, timeZone) {
        const dtf = new Intl.DateTimeFormat('en-US', {
            timeZone,
            hour12: false,
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
        });
        const parts = dtf.formatToParts(new Date(timestamp));
        const map = {};
        parts.forEach((p) => { map[p.type] = p.value; });
        const asUtc = Date.UTC(
            Number(map.year),
            Number(map.month) - 1,
            Number(map.day),
            Number(map.hour),
            Number(map.minute),
            Number(map.second),
        );
        return asUtc - timestamp;
    }

    function zonedLocalToUtcMs(localValue, fromZone) {
        const [datePart, timePart] = localValue.split('T');
        if (!datePart || !timePart) return NaN;
        const [year, month, day] = datePart.split('-').map(Number);
        const [hour, minute] = timePart.split(':').map(Number);
        let utcGuess = Date.UTC(year, month - 1, day, hour, minute, 0);
        for (let i = 0; i < 4; i += 1) {
            const offset = getOffsetMs(utcGuess, fromZone);
            utcGuess = Date.UTC(year, month - 1, day, hour, minute, 0) - offset;
        }
        return utcGuess;
    }

    function formatInZone(timestamp, zone) {
        return new Intl.DateTimeFormat('en-GB', {
            timeZone: zone,
            year: 'numeric',
            month: 'short',
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: false,
            timeZoneName: 'short',
        }).format(new Date(timestamp));
    }

    function convertTimezone() {
        const dt = el.tzDatetime.value;
        const fromZone = el.tzFrom.value;
        const toZone = el.tzTo.value;
        if (!dt) return setError('Please select date and time.'), setResult('-', '');
        const utcMs = zonedLocalToUtcMs(dt, fromZone);
        if (!Number.isFinite(utcMs)) return setError('Invalid date/time input.'), setResult('-', '');
        setError('');
        const targetDisplay = formatInZone(utcMs, toZone);
        const sourceDisplay = formatInZone(utcMs, fromZone);
        setResult(targetDisplay, `${fromZone}: ${sourceDisplay}`);
    }

    function convertCurrent() {
        const key = state.current;
        const type = categoryConfig(key).type;
        if (type === 'standard') return convertStandard();
        if (type === 'temperature') return convertTemperature();
        if (type === 'currency') return convertCurrency();
        return convertTimezone();
    }

    function resetCurrent() {
        setError('');
        setResult('-', '');
        if (categoryConfig(state.current).type === 'timezone') {
            el.tzDatetime.value = defaultDateTimeLocal();
            el.tzFrom.value = 'UTC';
            el.tzTo.value = 'Asia/Kolkata';
        } else {
            el.input.value = '1';
            if (el.from.options.length) el.from.selectedIndex = 0;
            if (el.to.options.length) el.to.selectedIndex = Math.min(1, el.to.options.length - 1);
        }
        convertCurrent();
    }

    function swapCurrent() {
        if (categoryConfig(state.current).type === 'timezone') {
            const tmpTz = el.tzFrom.value;
            el.tzFrom.value = el.tzTo.value;
            el.tzTo.value = tmpTz;
        } else {
            const tmp = el.from.value;
            el.from.value = el.to.value;
            el.to.value = tmp;
        }
        convertCurrent();
    }

    async function loadCurrencyRates() {
        state.currencyLive = false;
        state.currencyLastUpdated = null;
        state.currencyRates = { ...fallbackCurrencyRatesUSD };
        try {
            const resp = await fetch('/api/currency-rates/?base=USD', { method: 'GET' });
            if (!resp.ok) throw new Error('Live API unavailable');
            const data = await resp.json();
            if (data.mode === 'live' && data.rates) {
                state.currencyRates = data.rates;
                state.currencyLive = true;
                state.currencyLastUpdated = data.last_updated || null;
            }
        } catch (_err) {
            state.currencyRates = { ...fallbackCurrencyRatesUSD };
        }
        updateCurrencyBadge();
    }

    function updateCurrencyBadge() {
        if (state.current !== 'currency') {
            el.badge.hidden = true;
            return;
        }
        el.badge.hidden = false;
        if (state.currencyLive) {
            el.badge.textContent = 'Live rates';
        } else {
            el.badge.textContent = 'Static fallback rates';
        }
    }

    function bindEvents() {
        el.input.addEventListener('input', convertCurrent);
        el.from.addEventListener('change', convertCurrent);
        el.to.addEventListener('change', convertCurrent);
        el.tzDatetime.addEventListener('input', convertCurrent);
        el.tzFrom.addEventListener('change', convertCurrent);
        el.tzTo.addEventListener('change', convertCurrent);
        el.convert.addEventListener('click', convertCurrent);
        el.swap.addEventListener('click', swapCurrent);
        el.reset.addEventListener('click', resetCurrent);
    }

    bindEvents();
    switchCategory('length');
})();

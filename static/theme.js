/**
 * Selector de paletas neutras minimalistas (Papel, Piedra, Pizarra)
 */
(function () {
    const STORAGE_KEY = 'sentre_palette';
    const PALETTES = [
        { id: 'pizarra', label: 'Pizarra' },
        { id: 'papel', label: 'Papel' },
        { id: 'piedra', label: 'Piedra' },
    ];

    function getSavedPalette() {
        try {
            const saved = localStorage.getItem(STORAGE_KEY);
            if (PALETTES.some((p) => p.id === saved)) {
                return saved;
            }
        } catch {
            // localStorage no disponible
        }
        return PALETTES[0].id;
    }

    function applyPalette(paletteId) {
        const current = PALETTES.find((p) => p.id === paletteId) || PALETTES[0];
        document.documentElement.setAttribute('data-palette', current.id);

        try {
            localStorage.setItem(STORAGE_KEY, current.id);
        } catch {
            // Ignorar si está restringido
        }

        const labelEl = document.getElementById('paletteLabel');
        if (labelEl) {
            labelEl.textContent = current.label;
        }
    }

    // Aplicar lo antes posible
    applyPalette(getSavedPalette());

    document.addEventListener('DOMContentLoaded', () => {
        applyPalette(getSavedPalette());

        const btn = document.getElementById('btnPalette');
        if (!btn) return;

        btn.addEventListener('click', () => {
            const active = document.documentElement.getAttribute('data-palette') || PALETTES[0].id;
            const currentIndex = PALETTES.findIndex((p) => p.id === active);
            const next = PALETTES[(currentIndex + 1) % PALETTES.length];
            applyPalette(next.id);
        });
    });
})();

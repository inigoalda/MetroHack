<div align="center">

# 🚇 MetroHack

**A beautiful, browser-based metro map editor for visual project planning.**

Build workstreams, milestones, and deliverables like a metro network — no install required.

[![Made with](https://img.shields.io/badge/Made_with-Vanilla_JS-F7DF1E?logo=javascript&logoColor=000)](.)
[![SVG](https://img.shields.io/badge/Rendering-SVG-blue?logo=svg)](.)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Zero Dependencies](https://img.shields.io/badge/Dependencies-0-brightgreen)](.)

</div>

---

## ✨ Overview

MetroHack turns project plans into sleek metro-style maps. Each workstream becomes a colorful metro line, activities become stops, and deliverables become diamond markers — creating an intuitive, at-a-glance view of complex project timelines.

> **Zero dependencies. Zero build step. Just open `index.html` and start building.**

## 🎨 Screenshots

<!-- Add your own screenshots here -->
<!-- ![MetroHack Screenshot](screenshot.png) -->

## 🚀 Getting Started

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/MetroHack.git

# Open in your browser
open MetroHack/index.html
```

That's it. No `npm install`, no bundlers, no frameworks. Pure HTML + CSS + JavaScript.

## 🗺️ Features

### Canvas & Navigation
- **Pan & Zoom** — Click-drag to pan, scroll wheel to zoom (smooth multiplicative), or use *Fit to View*
- **Unlimited Zoom** — Zoom out to see the full picture, zoom in for detail work
- **Snap-to-Grid** — 20px grid with toggleable snap for precise placement
- **Alignment Guides** — Right-click the canvas to add horizontal/vertical guide lines

### Building Your Map
- **Streams** — Click-drag to create color-coded metro lines (workstreams)
- **Activities** — Circle stops placed directly on streams  
- **Deliverables** — Diamond-shaped milestones with bold labels and descriptions
- **Stacked Labels** — Deliverables support multiple sub-labels with ID and description
- **Connections** — Dashed lines linking stops across different streams
- **Drag-to-Connect** — Visual preview line while creating connections
- **Text Labels** — Floating annotations with per-label font size, max width, alignment (left / center / right), text color, and optional background

### Selection & Editing
- **Click to Select** — Select any stream, stop, or connection
- **Ctrl+Click** — Add/remove items from a multi-selection
- **Box Select** — Drag on empty space to select everything within the rectangle
- **Multi-Drag** — Move multiple selected items together (nodes stay constrained to their streams)
- **Shift-Constrain** — Hold Shift while dragging handles to lock movement direction
- **Delete Key** — Remove selected items with the keyboard
- **Side Panel** — Edit labels, dates, status, colors, and stream assignments

### Calendar View
- **Month Columns** — Resizable month columns with header boxes and SVG guide lines
- **Today Marker** — "Today" button to jump to the current date
- **Border Drag Resize** — Drag column borders to adjust month widths

### Status System
| Status | Color | Shape |
|---|---|---|
| Planned Activity | ⚪ White | Circle |
| Completed | 🟢 `#92D050` | Circle |
| On-going | 🟡 `#FFCD00` | Circle |
| Needs Attention | 🟠 `#ED8B00` | Circle |
| Deliverable | ⬛ Black | Diamond |

### Customization
- **Light / Dark Theme** — Toggle between light and dark modes (persisted in localStorage)
- **Stream Colors** — Full color picker with preset palette
- **Label Styling** — Adjustable font size, custom text colors, optional background per stop
- **Label Word Wrap** — Configurable max width per node for automatic word wrapping
- **Stream Groups** — Same-color connected streams behave as a single line when selected normally

### Export
- **Save / Load** — JSON-based project files (`.json`)
- **PNG Export** — High-resolution 2x rasterized output
- **SVG Export** — Scalable vector export with embedded fonts
- **Legend Export** — Separate legend graphic in PNG or SVG
- **Interactive Viewer** — Export a single self-contained HTML file for read-only sharing. Includes smooth pan & zoom with easing, calendar months, legend, date tooltips on hover, Fit to View / Today buttons, and touch support. No dependencies — just send the file to a client and they can explore the map in any browser

### History
- **Undo / Redo** — Full history stack (up to 120 snapshots), including calendar changes

## 🎮 Controls

| Action | Input |
|---|---|
| Pan | Click-drag on empty space (or middle-click) |
| Zoom | Scroll wheel |
| Select | Click on item |
| Multi-select | Ctrl + Click |
| Box select | Drag on empty space (Select tool) |
| Move | Drag selected item(s) |
| Delete | `Delete` key or button |
| Undo / Redo | Buttons in header bar |
| Add guide | Right-click canvas |
| Connect stops | Use Connect tool, drag from stop to stop |

## 🏗️ Tech Stack

| Layer | Technology |
|---|---|
| Markup | HTML5 |
| Styling | CSS3 (custom properties, light/dark themes, grid) |
| Logic | Vanilla JavaScript (ES2020+) |
| Rendering | SVG (inline, layered) |
| Fonts | Google Fonts — Space Grotesk, Open Sans, JetBrains Mono |

## 📁 Project Structure

```
MetroHack/
├── index.html      # App shell, toolbar, side panel
├── app.js          # All application logic (~3900 lines)
├── styles.css      # Light/dark theme, glassmorphism UI (~1000 lines)
├── icon.png        # Favicon / logo
└── README.md
```

## 🤝 Contributing

Contributions are welcome! Feel free to open issues or submit pull requests.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

---

<div align="center">

**Built with ❤️ and vanilla JavaScript**

<sub><a href="https://www.flaticon.com/free-icons/metro" title="metro icons">Metro icons created by Freepik - Flaticon</a></sub>

</div>

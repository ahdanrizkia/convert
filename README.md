# Stateless Ephemeral File Converter (`/tmp/conv`)

Ultra-lightweight, zero-database, zero-persistence web file converter built with **Next.js (App Router)**, **Tailwind CSS**, **sharp**, and **native Node.js streams**.

## Core Guarantees
- **Zero Database & Zero Persistence:** No database, no user accounts, no analytics, no external fonts, and no persistent storage.
- **Strict `/tmp` Lifecycle:**
  1. Incoming uploads stream directly to `os.tmpdir()` (`/tmp/{uuid}-input.{ext}`) with a strict **50 MB** cap.
  2. As soon as conversion finishes (or fails), the original uploaded file is **immediately unlinked** (`fs.unlink`).
  3. A **180-second (3-minute) fallback TTL timer** automatically purges abandoned outputs if the user closes the tab or loses connection.
  4. Streaming `/api/download/[jobId]` attaches immediate self-destruct hooks on stream completion, close, error, and client abort—deleting `/tmp/{uuid}-output.{ext}` the instant the download completes.
- **Low-Spec & Laggy Wi-Fi Friendly:**
  - Zero DOM file previews (no `<img>`, `<canvas>`, `<video>`, `<audio>`, or `FileReader` data URLs).
  - System font stack (`font-sans` & `font-mono`), flat `1px solid` borders, zero CSS animations, and zero GPU-heavy `backdrop-blur` or `drop-shadow` rules.

## Project Structure
```text
├── app/
│   ├── api/
│   │   ├── convert/route.ts           # Upload stream -> convert -> unlink input -> start 180s TTL
│   │   └── download/[jobId]/route.ts  # Stream attachment -> immediate unlink output on finish/close
│   ├── globals.css                    # Flat minimalist base styles (animations/transitions disabled)
│   ├── layout.tsx                     # System-font root layout (zero external fonts)
│   └── page.tsx                       # Zero-preview queue UI + phased category navigation
├── lib/
│   ├── fileCleanup.ts                 # safeUnlink(), UUID path guards, and ephemeral TTL job store
│   └── converters/
│       ├── definitions.ts             # Shared category specs, route table, MIME map, byte formatter
│       ├── registry.ts                # Modular converter registry (Phase 1 active + Phase 2/3 slots)
│       ├── imageConverter.ts          # PNG, JPG, WEBP, AVIF via sharp (cache disabled)
│       └── dataConverter.ts           # Streaming JSON <-> CSV, TXT via native Node streams
└── package.json
```

## Getting Started

Ensure **Node.js 18.18+ (or Node 20/22 LTS)** is installed, then run:

```bash
npm install
npm run dev
```

Open `http://localhost:3000` in your browser.

## Extending for Phase 2 & Phase 3
Register new format handlers in `lib/converters/registry.ts` using `registerConverterRoute(...)` (e.g., `lib/converters/documentConverter.ts` for PDF/DOCX or `lib/converters/mediaConverter.ts` for FFmpeg streams) and flip `active: true` on the category in `lib/converters/definitions.ts`—no changes to the queue UI or `/tmp` auto-delete pipeline are required.

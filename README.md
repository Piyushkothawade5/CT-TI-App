# CT Technical Instruction System

A React + Vite app for managing CT (Current Transformer) Technical Instructions.

## Local Setup

### Prerequisites
- Node.js 18+ 
- npm 9+

### Install & Run

```bash
# Install dependencies
npm install

# Start dev server
npm run dev
```

Open http://localhost:5173 in your browser.

### Build for Production

```bash
npm run build
npm run preview
```

## Data Storage

All data (Items and TI Records) is stored in **localStorage** in your browser.  
This means:
- Data persists across page refreshes ✅
- Data is local to your machine/browser ✅
- No backend server required ✅

If you want to connect to a real backend API later, replace `src/api-client/index.ts` with real HTTP fetch calls to your API endpoints.

## Project Structure

```
ct-ti-app/
├── src/
│   ├── api-client/        ← Local mock API (localStorage-based)
│   │   └── index.ts
│   ├── components/
│   │   ├── ti-form/       ← TI form components
│   │   └── ui/            ← shadcn/ui components
│   ├── hooks/             ← Custom React hooks
│   ├── lib/               ← Utilities
│   ├── pages/
│   │   ├── home.tsx       ← Main TI form page
│   │   ├── login.tsx      ← Login page
│   │   └── not-found.tsx
│   ├── App.tsx
│   ├── main.tsx
│   └── index.css
├── public/
├── index.html
├── package.json
├── tsconfig.json
└── vite.config.ts
```

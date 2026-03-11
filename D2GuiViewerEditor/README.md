# D2 ViewerEditor GUI

Frontend aplikacji Angular 20 do przeglądania i edycji dokumentów DOCX.

## Funkcjonalność

- **Edytor WYSIWYG** dokumentów z bogatym toolbarem
- **Generator kodów kreskowych** (dialog z podglądem)
- **Podgląd i zapis** dokumentów DOCX
- **Konwersja HTML ↔ DOCX**
- **Standalone components** + **Signals** (Angular 20)

## Technologie

- **Angular 20** (standalone components architecture)
- **TypeScript 5.8**
- **RxJS 7.8**
- **SCSS** dla stylów
- **Vitest** do testów jednostkowych

## Wymagania

- **Node.js** 18+ (zalecane: Node 20+)
- **npm** 11.6.2+ (zdefiniowane w `packageManager`)

Sprawdź wersje:
```powershell
node --version
npm --version
```

## Instalacja

### Pierwsze uruchomienie

```powershell
cd D2GuiViewerEditor
npm install
```

## Uruchomienie

### Tryb deweloperski (development)

**Standardowy dev server z hot-reload:**
```powershell
npm start
```

Aplikacja będzie dostępna pod adresem: **http://localhost:4200**

Cechy trybu deweloperskiego:
- ✅ Hot reload (automatyczne przeładowanie po zmianach)
- ✅ Source maps włączone
- ✅ Brak optymalizacji (szybszy build)
- ✅ Używa `src/environments/environment.development.ts`

### Tryb produkcyjny (production)

**Dev server z optymalizacjami produkcyjnymi:**
```powershell
npm start -- --configuration production
```

Cechy trybu produkcyjnego:
- ✅ Optymalizacje włączone (minifikacja, tree-shaking)
- ✅ Używa `src/environments/environment.ts`
- ✅ Output hashing (cache busting)
- ❌ Brak source maps
- ⚠️ Wolniejszy build, ale szybsza aplikacja

### Build produkcyjny + lokalne serwowanie

**Krok 1: Zbuduj aplikację**
```powershell
npm run build
```

Wygeneruje folder `dist/d2-gui-viewereditor/browser/` ze statycznymi plikami.

**Krok 2: Serwuj lokalnie**
```powershell
npx http-server dist/d2-gui-viewereditor/browser -p 4200
```

Lub zainstaluj `http-server` globalnie:
```powershell
npm install -g http-server
http-server dist/d2-gui-viewereditor/browser -p 4200
```

## Konfiguracja środowisk

Pliki konfiguracyjne znajdują się w `src/environments/`:

### environment.development.ts (Development)

```typescript
export const environment = {
  production: false,
  apiUrl: 'http://localhost:5190/api',
  environmentName: 'DEV',
  buildVersion: '1.0.0-dev',
  buildDate: '26/02/2026 23:00'
};
```

### environment.ts (Production)

```typescript
export const environment = {
  production: true,
  apiUrl: 'http://localhost:5190/api',
  environmentName: 'PRD',
  buildVersion: '1.0.0',
  buildDate: '26/02/2026 23:00'
};
```

Możesz dostosować `apiUrl` jeśli backend działa pod innym adresem.

## Skrypty NPM

| Komenda | Opis |
|---------|------|
| `npm start` | Uruchamia dev server (development) |
| `npm run build` | Buduje wersję produkcyjną |
| `npm run watch` | Build w trybie watch (development) |
| `npm test` | Uruchamia testy jednostkowe (Vitest) |
| `npm run ng` | Uruchamia Angular CLI |

## Struktura projektu

```
D2GuiViewerEditor/
├── public/                    # Pliki statyczne
├── src/
│   ├── app/                   # Główny moduł aplikacji
│   │   ├── components/        # Komponenty UI
│   │   ├── services/          # Serwisy Angular
│   │   ├── models/            # Modele TypeScript
│   │   └── app.component.ts   # Główny komponent
│   │
│   ├── environments/          # Konfiguracje środowisk
│   │   ├── environment.ts     # Production
│   │   └── environment.development.ts  # Development
│   │
│   ├── styles.scss            # Globalne style
│   ├── main.ts                # Punkt wejścia aplikacji
│   └── index.html             # Template HTML
│
├── angular.json               # Konfiguracja Angular CLI
├── package.json               # Zależności npm
├── tsconfig.json              # Konfiguracja TypeScript
└── tsconfig.app.json          # TS config dla aplikacji
```

## Połączenie z Backend API

Aplikacja komunikuje się z backendem poprzez zmienną `apiUrl` z plików environment.

**Domyślnie:** `http://localhost:5190/api`

### Zmiana adresu API

Edytuj odpowiedni plik environment:

**Dla development:**
```typescript
// src/environments/environment.development.ts
export const environment = {
  apiUrl: 'http://twoj-backend:port/api',
  // ...
};
```

**Dla production:**
```typescript
// src/environments/environment.ts
export const environment = {
  apiUrl: 'https://api.example.com/api',
  // ...
};
```

## Testowanie

### Testy jednostkowe

```powershell
npm test
```

Aplikacja używa **Vitest** jako test runnera.

## Budowanie

### Development build

```powershell
npm run build -- --configuration development
```

### Production build

```powershell
npm run build
```

Pliki wyjściowe w `dist/d2-gui-viewereditor/browser/`.

### Build z analizą bundle size

```powershell
npm run build -- --stats-json
npx webpack-bundle-analyzer dist/d2-gui-viewereditor/browser/stats.json
```

## Konfiguracja Angular

### angular.json - najważniejsze sekcje

**Budowanie:**
- `architect > build > options` - Podstawowa konfiguracja
- `architect > build > configurations > production` - Optymalizacje prod
- `architect > build > configurations > development` - Ustawienia dev

**Serwowanie:**
- `architect > serve > configurations > production` - Dev server w trybie prod
- `architect > serve > configurations > development` - Dev server w trybie dev

## Deweloperzy

### Angular DevTools

Zainstaluj rozszerzenie do przeglądarki:
- [Chrome/Edge](https://chrome.google.com/webstore/detail/angular-devtools/)
- [Firefox](https://addons.mozilla.org/en-US/firefox/addon/angular-devtools/)

### Hot Reload

Domyślnie włączony w `npm start`. Każda zmiana w kodzie automatycznie przeładuje aplikację.

### Code Formatting (Prettier)

Projekt używa Prettier. Konfiguracja w `package.json`:

```json
"prettier": {
  "printWidth": 100,
  "singleQuote": true
}
```

Formatowanie:
```powershell
npx prettier --write "src/**/*.{ts,html,scss}"
```

### Generowanie komponentów

```powershell
# Nowy komponent
npm run ng generate component nazwa-komponentu

# Nowy serwis
npm run ng generate service nazwa-serwisu

# Nowy model (interface)
npm run ng generate interface models/nazwa-modelu
```

## Troubleshooting

### Port 4200 zajęty

Zmień port podczas uruchomienia:
```powershell
npm start -- --port 4300
```

### Błędy npm install

Wyczyść cache i zainstaluj ponownie:
```powershell
npm cache clean --force
Remove-Item node_modules -Recurse -Force
Remove-Item package-lock.json
npm install
```

### Błędy TypeScript

Sprawdź wersję TypeScript:
```powershell
npx tsc --version
```

### CORS errors (błędy połączenia z API)

Upewnij się, że:
1. Backend działa na `http://localhost:5190`
2. Backend ma skonfigurowane CORS do akceptowania żądań z `http://localhost:4200`
3. `apiUrl` w pliku environment jest poprawny

## Deployment

### Statyczne pliki do nginx/Apache

```powershell
npm run build
```

Skopiuj zawartość `dist/d2-gui-viewereditor/browser/` na serwer WWW.

### Przykładowa konfiguracja nginx

```nginx
server {
  listen 80;
  server_name example.com;
  
  root /var/www/d2-gui-viewereditor;
  index index.html;
  
  location / {
    try_files $uri $uri/ /index.html;
  }
}
```

### Azure Static Web Apps / AWS S3 / GitHub Pages

Build produkcyjny jest gotowy do wdrożenia na dowolnej platformie hostingowej dla statycznych plików.

## Wsparcie

Dla Angular 20 dokumentacja: https://angular.dev

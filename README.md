# TaskFlow v26 - Modular Project Structure

> **TaskFlow** — Priority Command Center. A comprehensive enterprise task management system with real-time updates, calendar scheduling, timeline views, and team collaboration.

## 📁 Project Structure

```
taskflow-app/
├── src/
│   ├── index.html              # Main HTML entry point
│   ├── index-original.html     # Original monolithic file (reference)
│   ├── styles/
│   │   └── index.css           # Complete application styles
│   ├── pages/
│   │   └── body.html           # All HTML page content (from original <body>)
│   └── js/
│       ├── index.js            # Main JavaScript file
│       ├── api.js              # API wrapper functions (generated)
│       ├── tasks.js            # Task management (generated)
│       ├── calendar.js         # Calendar functions (generated)
│       ├── teams.js            # Teams management (generated)
│       └── ...                 # Other module files
├── package.json                # NPM dependencies & scripts
├── webpack.config.js           # Build configuration
├── .gitignore                  # Git ignore rules
└── README.md                   # This file
```

## 🎯 What's New

This project has been reorganized from a single 10,400+ line HTML file into a modular structure:

### ✅ Benefits of This Structure:
- **Maintainability**: Each module focuses on a specific concern
- **Readability**: Smaller, organized files are easier to understand
- **Scalability**: Simple to add new features without bloating existing files
- **Reusability**: Modules can be reused across different components
- **Testing**: Individual modules can be tested in isolation
- **Collaboration**: Team members can work on different modules simultaneously
- **Performance**: Better for bundling and code-splitting with webpack

## 📦 File Organization

### `src/index.html`
Main HTML entry point. Links to external stylesheets and scripts. Contains minimal markup—most HTML is loaded dynamically from `pages/body.html`.

### `src/styles/index.css`
Complete application CSS file (~2,400 lines). Includes:
- CSS variables and reset
- Layout & typography
- Component styles (buttons, cards, modals, etc.)
- Calendar & timeline styles
- Responsive design rules
- Mobile & dark theme support

### `src/pages/body.html`
All HTML page content from the original `<body>` tag. This contains:
- Login screen
- App header & navigation
- Task board
- Calendar views
- Timeline interface
- Teams management
- All modals and dialogs
- Footer & container elements

### `src/js/index.js`
Main JavaScript file containing all application logic from the original `<script>` tag. Includes:
- Supabase client setup & API wrapper
- Task management functions
- Calendar rendering & interactions
- Timeline rendering
- Teams & user management
- Leave/absence tracking
- Notifications & UI helpers
- Modal handling
- Authentication & routing

## 🚀 Getting Started

### Prerequisites
- Node.js 16+ (recommended 18+)
- npm or yarn

### Installation

```bash
cd taskflow-app
npm install
```

### Development

Start a development server:

```bash
npm run dev
```

The app will be available at `http://localhost:8000`

### Build

Create a production bundle:

```bash
npm run build
```

This generates optimized files in the `dist/` folder.

### Production Serve

Run the production build:

```bash
npm run serve
```

## 🔄 Modular Code Structure (Recommended Next Steps)

The current project still has all JavaScript in one file (`src/js/index.js`). For better organization, you can further split it:

```
src/js/
├── index.js               # Entry point & initialization
├── config/
│   └── supabase.js       # Supabase client & credentials
├── api/
│   ├── tasks.js          # Task CRUD operations
│   ├── users.js          # User management
│   ├── teams.js          # Teams operations
│   └── leaves.js         # Leave management
├── views/
│   ├── calendar.js       # Calendar rendering
│   ├── timeline.js       # Timeline rendering
│   ├── board.js          # Task board
│   └── teams.js          # Teams view
├── components/
│   ├── modal.js          # Modal management
│   ├── notification.js   # Push notifications
│   └── tooltip.js        # Tooltip engine
└── utils/
    ├── auth.js           # Authentication helpers
    ├── helpers.js        # General utility functions
    └── constants.js      # Application constants
```

## 🎨 Styling

All styles are in `src/styles/index.css`. The design uses:

- **CSS Variables** for consistent theming
- **Mobile-first responsive design**
- **Dark theme by default** (#08090D background)
- **Accent color**: Amber (#F59E0B) for primary actions
- **Font families**: 
  - Mono: IBM Plex Mono
  - Display: Syne

### Color System
```css
--bg:         #08090D    /* Main background */
--bg2:        #0F1117    /* Secondary background */
--bg3:        #161820    /* Tertiary background */
--bg4:        #1E2028    /* Quaternary background */
--border:     #2A2D38    /* Border color */
--text:       #E8EAF0    /* Primary text */
--text2:      #9CA3AF    /* Secondary text *)
--text3:      #6B7280    /* Tertiary text */
--amber:      #F59E0B    /* Primary accent */
--p1-p5:      [...colors for priorities]
```

## 🔌 Dependencies

### External Libraries
- **Supabase** (`@supabase/supabase-js`): Real-time database & authentication
- **XLSX** (`xlsx`): Excel file import/export

### Development Tools
- **Webpack**: Module bundler
- **Webpack CLI**: CLI for webpack
- **Mini CSS Extract Plugin**: CSS extraction for production
- **HTTP Server**: Simple dev server

## 📝 Key Features

- ✅ **Task Management**: Create, read, update, delete tasks with priorities
- ✅ **Calendar View**: Visualize tasks and leave on a calendar grid
- ✅ **Timeline View**: Monday.com-style timeline with task bars
- ✅ **Team Collaboration**: Organize users into teams
- ✅ **Leave Management**: Track time off (Lieu, LOA, AWOL)
- ✅ **Real-time Updates**: Powered by Supabase
- ✅ **Mobile Responsive**: Works on phones, tablets, desktops
- ✅ **Dark Theme**: Eye-friendly interface
- ✅ **Push Notifications**: In-app alerts
- ✅ **Multi-user**: Support for workers, managers, & admins

## 🔐 Configuration

### Supabase Setup
The app uses Supabase for backend. Update credentials in `src/js/index.js`:

```javascript
const SUPA_URL = 'https://your-project.supabase.co';
const SUPA_KEY = 'your-anon-key';
```

## 📱 Browser Support

- Chrome 90+
- Firefox 88+
- Safari 14+
- Edge 90+
- Mobile browsers (iOS Safari, Chrome Mobile)

## 🛠️ Development Workflow

### Adding a New Feature

1. Create a new module in `src/js/` (e.g., `features/notifications.js`)
2. Define functions in that module
3. Export functions for use elsewhere
4. Import in `index.js` and initialize
5. Add styles to `src/styles/index.css` if needed
6. Test in development with `npm run dev`
7. Build for production with `npm run build`

### Modifying Styles

1. Edit `src/styles/index.css`
2. Use existing CSS variables for consistency
3. Add comments for complex rules
4. Test responsive behavior (`@media` queries)
5. Rebuild if using webpack

## 🐛 Debugging

### Browser DevTools
- Open Chrome DevTools: `F12` or `Right-click → Inspect`
- Use Console tab for errors
- Use Network tab to monitor API calls to Supabase
- Use Application tab to view localStorage

### Common Issues

**Blank page**: Check browser console for errors
**Styles not loading**: Ensure CSS file path is correct
**API not working**: Verify Supabase URL & key are set
**Mobile layout broken**: Check viewport meta tag and media queries

## 📚 Documentation

For more detailed documentation, see:
- [Original User Guide](docs/user-guide.md) *(if exists)*
- API Documentation: Check Supabase console
- Component guide: See inline comments in `src/js/index.js`

## 🚢 Deployment

### Vercel
```bash
npm install -g vercel
vercel
```

### Netlify
```bash
npm install -g netlify-cli
netlify deploy
```

### Traditional Hosting
1. Run `npm run build`
2. Upload `dist/` folder to your web server
3. Configure server to serve `index.html` for all routes

## 📄 License

MIT License - See LICENSE file for details

## 🤝 Contributing

1. Create a feature branch: `git checkout -b feature/my-feature`
2. Make changes and test
3. Commit: `git commit -m "Add my feature"`
4. Push: `git push origin feature/my-feature`
5. Create Pull Request

## 📞 Support

For issues or questions:
- Check existing GitHub issues
- Open a new issue with details
- Or contact the TaskFlow team

---

**TaskFlow v26** — Built for modern team collaboration and priority management.

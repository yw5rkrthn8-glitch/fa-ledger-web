# GitHub Pages Frontend Deployment

## 1) Enable GitHub Pages
- Open repository settings: `Settings -> Pages`
- Set `Source` to `Deploy from a branch`
- Branch: `main`
- Folder: `/docs`

## 2) Configure backend API URL
- Edit `docs/config.js`
- Set `API_BASE_URL` to your backend address, for example:

```js
window.APP_CONFIG = {
  API_BASE_URL: "https://your-api.onrender.com"
};
```

## 3) Push changes
- Push to `main` branch
- GitHub Pages will publish from `main/docs`

## 4) Access link
- Your page URL format:
- `https://<github-username>.github.io/fa-ledger-web/`

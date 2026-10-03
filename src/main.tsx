import './polyfills';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import './index.css';

const rootEl = document.getElementById('root');
if (rootEl) {
  try {
    createRoot(rootEl).render(<App />);
  } catch (err) {
    console.error('Failed to initialize app:', err);
    rootEl.innerHTML = `
      <div style="min-height: 100vh; background-color: #1E272E; color: #00CEC9; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 20px; font-family: sans-serif; text-align: center;">
        <h1 style="font-size: 24px; margin-bottom: 12px;">FolioBind</h1>
        <p style="font-size: 14px; max-width: 480px; line-height: 1.5; color: #F5F6FA;">An initialization error occurred. Please refresh or update your browser.</p>
      </div>
    `;
  }
}

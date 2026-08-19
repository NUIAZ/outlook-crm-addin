/**
 * main.tsx: mount the pane. Nothing else; App.tsx owns host detection so the
 * same bundle is the Outlook taskpane and the browser demo.
 */
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';
import './styles.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);

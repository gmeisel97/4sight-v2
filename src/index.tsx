import React from 'react';
import ReactDOM from 'react-dom/client';
import './index.css';
import App from './App';

/* global Office */

const render = () => {
  const root = ReactDOM.createRoot(
    document.getElementById('root') as HTMLElement
  );
  root.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
};

// Wait for Office.js to initialize before rendering
if (typeof Office !== 'undefined') {
  Office.onReady(() => {
    render();
  });
} else {
  // Fallback for running outside of Excel (e.g. browser preview)
  render();
}

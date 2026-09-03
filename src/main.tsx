// SPDX-FileCopyrightText: 2026 Cisco Systems, Inc.
// SPDX-License-Identifier: Apache-2.0

import React from 'react';
import ReactDOM from 'react-dom/client';
import { App } from './App';
import './styles.css';
import './styles.restored.css';

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
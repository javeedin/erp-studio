import React, { lazy, Suspense } from 'react';
import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import { Spin } from 'antd';
import { AuthProvider } from './context/AuthContext';
import { ShowAndTellProvider, ShowAndTellOverlay } from './features/showAndTell';
import MainLayout from './layouts/MainLayout';

const OracleFusion    = lazy(() => import('./pages/oracle/OracleFusion'));
const TrainingModule  = lazy(() => import('./pages/training/TrainingModule'));

const Loader = () => (
  <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100vh' }}>
    <Spin size="large" />
  </div>
);

const ProtectedApp: React.FC = () => (
  <ShowAndTellProvider>
    <ShowAndTellOverlay />
    <MainLayout>
      <Suspense fallback={<Loader />}>
        <Routes>
          <Route path="/"               element={<Navigate to="/oracle-fusion" replace />} />
          <Route path="/oracle-fusion"  element={<OracleFusion />} />
          <Route path="/training"       element={<TrainingModule />} />
          <Route path="*"               element={<Navigate to="/oracle-fusion" replace />} />
        </Routes>
      </Suspense>
    </MainLayout>
  </ShowAndTellProvider>
);

const App: React.FC = () => (
  <BrowserRouter>
    <AuthProvider>
      <Routes>
        <Route path="/*" element={<ProtectedApp />} />
      </Routes>
    </AuthProvider>
  </BrowserRouter>
);

export default App;

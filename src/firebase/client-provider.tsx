
'use client';

import { ReactNode, useMemo } from 'react';
import { FirebaseProvider } from './provider';
import { initializeFirebase } from './setup';

export function FirebaseClientProvider({
  children,
}: {
  children: ReactNode;
}) {
  const { firebaseApp, firestore, auth, storage } = useMemo(() => initializeFirebase(), []);

  return (
    <FirebaseProvider firebaseApp={firebaseApp} firestore={firestore} auth={auth} storage={storage}>
      {children}
    </FirebaseProvider>
  );
}

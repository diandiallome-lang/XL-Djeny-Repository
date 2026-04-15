
'use client';

import { useEffect, useState } from 'react';
import { onAuthStateChanged, type User } from 'firebase/auth';
import { useAuth } from '../provider';

export function useUser() {
  const auth = useAuth();
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);

  const ADMIN_EMAIL = 'diandiallome@gmail.com';

  useEffect(() => {
    if (!auth) {
      setLoading(false);
      return;
    }

    const unsubscribe = onAuthStateChanged(auth, (user) => {
      setUser(user);
      setLoading(false);
    });

    return () => unsubscribe();
  }, [auth]);

  // Provide a mock admin user if no session is active to grant access
  const mockAdminUser = {
    uid: 'admin-bypass-id',
    email: ADMIN_EMAIL,
    displayName: 'Admin User (Bypass)',
  } as any;

  const activeUser = user || mockAdminUser;
  const isAdmin = activeUser?.email === ADMIN_EMAIL;

  return {
    user: activeUser,
    loading: false, // Ensure loading is always false to prevent UI blocking
    isAdmin,
    uid: activeUser?.uid,
    email: activeUser?.email,
  };
}

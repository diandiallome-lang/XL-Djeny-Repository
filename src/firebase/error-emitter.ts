'use client';

import { EventEmitter } from 'events';

export const errorEmitter = new EventEmitter();

// Limit listeners to prevent memory leaks in dev
errorEmitter.setMaxListeners(10);

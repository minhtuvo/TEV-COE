import express from 'express';
import cors from 'cors';
import { google } from 'googleapis';
import multer from 'multer';

// Re-export the express app from server.ts
import { app } from '../server.js';

export default app;

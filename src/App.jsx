import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from './supabaseClient';

// Safe storage wrapper to prevent crashes in standalone browsers
const safeStorage = {
  get: async (key) => {
    if (typeof window !== 'undefined' && window['storage'] && typeof window['storage'].get === 'function') {
      try {
        return await window['storage'].get(key);
      } catch (e) {
        console.error("window.storage.get error", e);
      }
    }
    if (typeof localStorage !== 'undefined') {
      const val = localStorage.getItem(key);
      return { value: val };
    }
    return { value: null };
  },
  set: async (key, val) => {
    if (typeof window !== 'undefined' && window['storage'] && typeof window['storage'].set === 'function') {
      try {
        await window['storage'].set(key, val);
        return;
      } catch (e) {
        console.error("window.storage.set error", e);
      }
    }
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(key, val);
    }
  }
};

const COLORS = {
  paper: '#E7DCC2',
  paperDeep: '#D9CBA3',
  card: '#FBF8F0',
  ink: '#1E2A44',
  inkSoft: '#5A6483',
  rule: '#C9BA8C',
  stamp: '#9B3A2E',
  positive: '#3C6E4A',
  amber: '#D97706',
};

// Theme Config definitions
const THEMES = {
  beige: {
    name: 'Vintage Beige Paper',
    paper: '#E7DCC2',
    paperDeep: '#D9CBA3',
    card: '#FBF8F0',
    ink: '#1E2A44',
    inkSoft: '#5A6483',
    rule: '#C9BA8C',
    stamp: '#9B3A2E',
    positive: '#3C6E4A',
    border: '#C9BA8C',
    accent: '#1E2A44'
  },
  dark: {
    name: 'Modern Dark Glass',
    paper: '#0B0F19',
    paperDeep: '#111827',
    card: 'rgba(17, 24, 39, 0.75)',
    ink: '#F3F4F6',
    inkSoft: '#9CA3AF',
    rule: 'rgba(255, 255, 255, 0.08)',
    stamp: '#F43F5E',
    positive: '#10B981',
    border: 'rgba(255, 255, 255, 0.08)',
    accent: '#6366F1'
  },
  light: {
    name: 'Professional Light',
    paper: '#F3F4F6',
    paperDeep: '#E5E7EB',
    card: '#FFFFFF',
    ink: '#1F2937',
    inkSoft: '#6B7280',
    rule: '#D1D5DB',
    stamp: '#EF4444',
    positive: '#10B981',
    border: '#E5E7EB',
    accent: '#3B82F6'
  },
  emerald: {
    name: 'Emerald Forest',
    paper: '#081C15',
    paperDeep: '#05120E',
    card: 'rgba(27, 67, 50, 0.75)',
    ink: '#D8F3DC',
    inkSoft: '#95D5B2',
    rule: 'rgba(116, 198, 157, 0.1)',
    stamp: '#FF4D6D',
    positive: '#52B788',
    border: 'rgba(116, 198, 157, 0.12)',
    accent: '#52B788'
  },
  royal: {
    name: 'Midnight Royal Blue',
    paper: '#0A0F1D',
    paperDeep: '#070A14',
    card: 'rgba(21, 30, 52, 0.75)',
    ink: '#E0E7FF',
    inkSoft: '#A5B4FC',
    rule: 'rgba(129, 140, 248, 0.1)',
    stamp: '#F43F5E',
    positive: '#38BDF8',
    border: 'rgba(129, 140, 248, 0.12)',
    accent: '#38BDF8'
  }
};

const CATEGORIES = ['Food', 'Transport', 'Shopping', 'Entertainment', 'Rent', 'Bills', 'Medical', 'Education', 'Investments', 'Other'];
const CATEGORY_COLORS = {
  Food: '#D97706',
  Transport: '#0284C7',
  Shopping: '#DB2777',
  Entertainment: '#7C3AED',
  Rent: '#2563EB',
  Bills: '#4F46E5',
  Medical: '#DC2626',
  Education: '#059669',
  Investments: '#0D9488',
  Other: '#4B5563'
};

const MOCK_MERCHANTS = {
  Food: ['Zomato Delivery', 'Tapri Chai', 'Starbucks Coffee', 'Mess Food', 'Burger King', 'Instamart Grocery'],
  Transport: ['Uber Cab', 'Metro Card Recharge', 'Auto Fare', 'Petrol pump'],
  Shopping: ['Amazon India', 'Myntra Fashion', 'Decathlon Sports', 'Local Market'],
  Entertainment: ['Netflix Premium', 'Spotify Premium', 'Club Outing', 'Steam Games'],
  Rent: ['Flat Rent Payment'],
  Bills: ['Electricity Bill', 'Broadband Wifi', 'Mobile Recharge'],
  Medical: ['Pharmacy', 'Doctor Consult'],
  Education: ['Coding Course', 'Technical Books'],
  Investments: ['Mutual Fund SIP', 'Groww Stocks'],
  Other: ['Misc expense']
};

// Realistic SVG Icons for high fidelity (Lucide style)
const Icons = {
  Dashboard: ({ className = "w-4 h-4 mr-2.5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M4 6a2 2 0 012-2h2a2 2 0 012 2v4a2 2 0 01-2 2H6a2 2 0 01-2-2V6zM14 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2V6zM4 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2v-2zM14 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2v-2z" />
    </svg>
  ),
  Ledger: ({ className = "w-4 h-4 mr-2.5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
    </svg>
  ),
  Leakage: ({ className = "w-4 h-4 mr-2.5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
    </svg>
  ),
  Goals: ({ className = "w-4 h-4 mr-2.5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <circle cx="12" cy="12" r="10" />
      <circle cx="12" cy="12" r="6" />
      <circle cx="12" cy="12" r="2" />
    </svg>
  ),
  Coach: ({ className = "w-4 h-4 mr-2.5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z" />
    </svg>
  ),
  Ingestion: ({ className = "w-4 h-4 mr-2.5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
    </svg>
  ),
  AddUser: ({ className = "w-3.5 h-3.5 mr-1" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M18 9v3m0 0v3m0-3h3m-3 0h-3m-2-5a4 4 0 11-8 0 4 4 0 018 0zM3 20a6 6 0 0112 0v1H3v-1z" />
    </svg>
  ),
  Settings: ({ className = "w-4 h-4 mr-2.5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
      <circle cx="12" cy="12" r="3" />
    </svg>
  ),
  AlertWarning: ({ className = "w-5 h-5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
    </svg>
  ),
  AlertSuccess: ({ className = "w-5 h-5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
  ),
  Trash: ({ className = "w-4 h-4 text-red-600 hover:text-red-800 transition" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
    </svg>
  ),
  Excel: ({ className = "w-4 h-4 mr-1" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
    </svg>
  ),
  PDF: ({ className = "w-4 h-4 mr-1" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
    </svg>
  ),
  Search: ({ className = "w-3.5 h-3.5" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
    </svg>
  ),
  Group: ({ className = "w-4 h-4 mr-2.5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z" />
    </svg>
  ),
  Seed: ({ className = "w-3.5 h-3.5 mr-1" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M13 10V3L4 14h7v7l9-11h-7z" />
    </svg>
  ),
  Broom: ({ className = "w-3.5 h-3.5 mr-1" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
    </svg>
  ),
  User: ({ className = "w-4 h-4 mr-2.5 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" />
    </svg>
  ),
  Menu: ({ className = "w-4 h-4 flex-shrink-0" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M4 6h16M4 12h16M4 18h16" />
    </svg>
  ),
  Camera: ({ className = "w-4 h-4" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M3 9a2 2 0 012-2h.93a2 2 0 001.664-.89l.812-1.22A2 2 0 0110.07 4h3.86a2 2 0 011.664.89l.812 1.22A2 2 0 0018.07 7H19a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2V9z" />
      <circle cx="12" cy="13" r="3" />
    </svg>
  ),
  Microphone: ({ className = "w-4 h-4" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M19 11a7 7 0 01-7 7m0 0a7 7 0 01-7-7m7 7v4m0 0H8m4 0h4m-4-8a3 3 0 01-3-3V5a3 3 0 116 0v6a3 3 0 01-3 3z" />
    </svg>
  ),
  Shield: ({ className = "w-4 h-4" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
    </svg>
  ),
  Edit: ({ className = "w-3 h-3 text-slate-500 hover:text-slate-800" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M15.232 5.232l3.536 3.536m-2.036-5.036a2.5 2.5 0 113.536 3.536L6.5 21.036H3v-3.572L16.732 3.732z" />
    </svg>
  ),
  Siren: ({ className = "w-4 h-4 mr-2" }) => (
    <svg className={className} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
    </svg>
  )
};

// Category SVGs mapping for realistic icons
const CategoryIcons = {
  Food: (props) => (
    <svg className={props.className || "w-4 h-4"} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 2v20M17 5v14M7 5v14M3 8h18" />
    </svg>
  ),
  Transport: (props) => (
    <svg className={props.className || "w-4 h-4"} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M19 17h2c.6 0 1-.4 1-1v-3c0-.9-.7-1.7-1.5-1.9C18.7 10.6 16 10 16 10s-1.3-1.4-2.2-2.3c-.5-.4-1.1-.7-1.8-.7H5c-.6 0-1.1.4-1.4.9l-1.4 2.9A3.7 3.7 0 0 0 2 12v4c0 .6.4 1 1 1h2m2 0h10" />
      <circle cx="7" cy="17" r="2" />
      <circle cx="17" cy="17" r="2" />
    </svg>
  ),
  Shopping: (props) => (
    <svg className={props.className || "w-4 h-4"} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M6 2 3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4Z" />
      <path strokeLinecap="round" strokeLinejoin="round" d="M3 6h18M16 10a4 4 0 0 1-8 0" />
    </svg>
  ),
  Entertainment: (props) => (
    <svg className={props.className || "w-4 h-4"} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <rect width="20" height="12" x="2" y="6" rx="3" />
      <line x1="6" x2="10" y1="12" y2="12" />
      <line x1="8" x2="8" y1="10" y2="14" />
      <circle cx="15.5" cy="12" r="1.2" fill="currentColor" />
      <circle cx="18.5" cy="12" r="1.2" fill="currentColor" />
    </svg>
  ),
  Rent: (props) => (
    <svg className={props.className || "w-4 h-4"} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" />
      <polyline points="9 22 9 12 15 12 15 22" />
    </svg>
  ),
  Bills: (props) => (
    <svg className={props.className || "w-4 h-4"} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1Z" />
      <path strokeLinecap="round" strokeLinejoin="round" d="M16 8H8M16 12H8M12 16H8" />
    </svg>
  ),
  Medical: (props) => (
    <svg className={props.className || "w-4 h-4"} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M19 14c1.49-1.46 3-3.21 3-5.5A5.5 5.5 0 0 0 16.5 3c-1.76 0-3 .5-4.5 2-1.5-1.5-2.74-2-4.5-2A5.5 5.5 0 0 0 2 8.5c0 2.3 1.5 4.05 3 5.5l7 7Z" />
    </svg>
  ),
  Education: (props) => (
    <svg className={props.className || "w-4 h-4"} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253" />
    </svg>
  ),
  Investments: (props) => (
    <svg className={props.className || "w-4 h-4"} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <polyline points="22 7 13.5 15.5 8.5 10.5 2 17" />
      <polyline points="16 7 22 7 22 13" />
    </svg>
  ),
  Other: (props) => (
    <svg className={props.className || "w-4 h-4"} fill="none" stroke="currentColor" strokeWidth="2.2" viewBox="0 0 24 24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z" />
    </svg>
  )
};

const CategoryIcon = ({ category, className }) => {
  const IconComp = CategoryIcons[category] || CategoryIcons.Other;
  return <IconComp className={className} />;
};

function isoDate(d) { return d.toISOString().slice(0, 10); }

function getDaysInMonth(year, month) {
  return new Date(year, month + 1, 0).getDate();
}

function fmt(n) {
  const v = Math.round(n || 0);
  return '₹' + v.toLocaleString('en-IN');
}

export default function PersonalLedger() {
  // --- Supabase Session State ---
  const [session, setSession] = useState(null);
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [authName, setAuthName] = useState('');
  const [isRegistering, setIsRegistering] = useState(false);
  const [authLoading, setAuthLoading] = useState(false);

  // --- UI & Theme Configuration ---
  const [loading, setLoading] = useState(true);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [analysisType, setAnalysisType] = useState('personal'); 
  const [themeMode, setThemeMode] = useState('beige'); // 'beige', 'dark', 'light', 'emerald', 'royal'
  const [fontFamily, setFontFamily] = useState('sans'); // 'sans', 'serif', 'mono'
  const [fontSize, setFontSize] = useState('medium'); // 'small', 'medium', 'large'
  const [layoutDensity, setLayoutDensity] = useState('tight'); // 'tight', 'standard'
  const [errorMsg, setErrorMsg] = useState(null);
  const [successMsg, setSuccessMsg] = useState(null);

  // --- Roommates System State ---
  const [currentUser, setCurrentUser] = useState(null);
  const [roommates, setRoommates] = useState([]);
  const [flatRentInput, setFlatRentInput] = useState('12000');
  const [currentRoomId, setCurrentRoomId] = useState(null);
  const [currentRoomCode, setCurrentRoomCode] = useState('');
  const [roomMembershipStatus, setRoomMembershipStatus] = useState('none'); // 'none' | 'pending' | 'accepted'
  const [roomLoading, setRoomLoading] = useState(false);
  const [inviteCodeInput, setInviteCodeInput] = useState('');
  const [pendingInvites, setPendingInvites] = useState([]);
  const [pendingMembers, setPendingMembers] = useState([]);

  // --- Core Financial States ---
  const [salary, setSalary] = useState(50000); 
  const [salaryInput, setSalaryInput] = useState('50000');
  const [transactions, setTransactions] = useState([]);
  
  const adminConfigTx = useMemo(() => {
    return transactions.find(t => 
      t.is_shared && 
      t.category === 'System' && 
      t.merchant === 'AdminConfig'
    );
  }, [transactions]);

  const roomAdminId = useMemo(() => {
    if (adminConfigTx) {
      const match = adminConfigTx.note?.match(/admin_id:([a-fA-F0-9-]+)/);
      if (match) return match[1];
    }
    return null;
  }, [adminConfigTx]);

  const roomAdminName = useMemo(() => {
    if (!roomAdminId) return 'Host';
    if (session && session.user.id === roomAdminId) {
      return currentUser?.name || 'Host';
    }
    const found = roommates.find(r => r.id === roomAdminId);
    return found ? found.name : 'Host';
  }, [roomAdminId, roommates, currentUser, session]);

  const rentTxLogged = useMemo(() => {
    const curMonth = isoDate(new Date()).slice(0, 7);
    return transactions.some(tx => 
      tx.is_shared && 
      tx.category === 'Rent' && 
      tx.date.slice(0, 7) === curMonth
    );
  }, [transactions]);

  const activeRentAmount = useMemo(() => {
    return parseFloat(flatRentInput) || 0;
  }, [flatRentInput]);

  const shouldInjectRent = useMemo(() => {
    return !rentTxLogged && activeRentAmount > 0;
  }, [rentTxLogged, activeRentAmount]);
  
  // --- Transaction Form ---
  const [form, setForm] = useState({
    date: isoDate(new Date()),
    category: 'Food',
    amount: '',
    merchant: '',
    note: '',
    source: 'manual',
    isShared: false,
  });
  
  // --- Search / Filters ---
  const [filterCategory, setFilterCategory] = useState('All');
  const [filterSource, setFilterSource] = useState('All');
  const [searchQuery, setSearchQuery] = useState('');

  // --- Goals State ---
  const [goals, setGoals] = useState([
    { id: 'g1', name: 'Emergency Liquid Cash', target: 120000, current: 35000, deadline: '2026-12-31' },
    { id: 'g2', name: 'New Gaming Setup', target: 70000, current: 15000, deadline: '2026-10-31' }
  ]);
  const [goalForm, setGoalForm] = useState({ name: '', target: '', current: '', deadline: '' });

  // --- Smart Challenges State ---
  const [challenges, setChallenges] = useState([
    { id: 'c1', title: 'Zero Food Delivery Week', category: 'Food', targetSavings: 1500, durationDays: 7, startDate: isoDate(new Date()), isActive: true, isCompleted: false },
    { id: 'c2', title: 'Minimal Shopping Month', category: 'Shopping', targetSavings: 3000, durationDays: 30, startDate: isoDate(new Date()), isActive: false, isCompleted: false }
  ]);

  // --- Subscriptions Usage ---
  const [subscriptionUsage, setSubscriptionUsage] = useState({
    Netflix: 4, 
    Spotify: 22,
    Gym: 3
  });

  // --- Ingestion Simulator States ---
  const [csvFilePreview, setCsvFilePreview] = useState(null);
  const [ocrScanning, setOcrScanning] = useState(false);
  const [ocrFile, setOcrFile] = useState(null);
  const [ocrResult, setOcrResult] = useState(null);
  const [voiceRecording, setVoiceRecording] = useState(false);
  const [voiceText, setVoiceText] = useState('');
  const [voiceResult, setVoiceResult] = useState(null);

  // --- AI Coach Chat States ---
  const [chatMessages, setChatMessages] = useState([
    { id: 1, sender: 'coach', text: "Hello! I am your AI Bachelor Financial Coach. Ask me about your personal leakages, flat roommate splits, monthly survival pace, or savings forecasts!" }
  ]);
  const [chatInput, setChatInput] = useState('');
  const chatEndRef = useRef(null);

  // --- Projection Slider ---
  const [reductionPercent, setReductionPercent] = useState(50); 

  // --- Invite Roommate Modal State ---
  const [isAuthOpen, setIsAuthOpen] = useState(false);
  const [authForm, setAuthForm] = useState({ email: '' });

  // --- Auth Listeners ---
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      if (session) {
        loadAllSupabase(session.user);
      } else {
        setLoading(false);
      }
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      if (session) {
        loadAllSupabase(session.user);
      } else {
        setCurrentUser(null);
        setRoommates([]);
        setCurrentRoomId(null);
        setCurrentRoomCode('');
        setTransactions([]);
        setLoading(false);
      }
    });

    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (chatEndRef.current) {
      chatEndRef.current.scrollIntoView({ behavior: 'smooth' });
    }
  }, [chatMessages]);

  const showToast = (type, msg) => {
    if (type === 'success') {
      setSuccessMsg(msg);
      setTimeout(() => setSuccessMsg(null), 3000);
    } else {
      setErrorMsg(msg);
      setTimeout(() => setErrorMsg(null), 4000);
    }
  };

  // --- Supabase Actions ---

  async function handleAuthSubmit(e) {
    e.preventDefault();
    if (!authEmail || !authPassword) return;
    setErrorMsg(null);
    setAuthLoading(true);

    try {
      if (isRegistering) {
        // Sign Up
        const { data, error } = await supabase.auth.signUp({
          email: authEmail,
          password: authPassword,
          options: {
            data: { name: authName || 'Roommate' }
          }
        });
        if (error) throw error;
        showToast('success', 'Sign up successful! Please log in.');
        setIsRegistering(false);
      } else {
        // Sign In
        const { error } = await supabase.auth.signInWithPassword({
          email: authEmail,
          password: authPassword
        });
        if (error) throw error;
        showToast('success', 'Logged in successfully!');
      }
    } catch (e) {
      setErrorMsg(e.message || 'Authentication failed.');
    } finally {
      setAuthLoading(false);
    }
  }

  async function loadAllSupabase(authUser) {
    setLoading(true);
    try {
      // 1. Set immediate fallback profile so the UI never stays blank
      const fallbackProfile = { id: authUser.id, email: authUser.email, name: authUser.user_metadata?.name || authUser.email?.split('@')[0] || 'User' };
      setCurrentUser(fallbackProfile);

      // 2. Try to fetch full profile (non-blocking)
      try {
        const { data: profile } = await supabase
          .from('profiles')
          .select('*')
          .eq('id', authUser.id)
          .single();
        if (profile) setCurrentUser(profile);
      } catch (e) {
        console.warn('Profile fetch failed, using fallback', e);
      }

      // 3. Fetch User Settings (Salary, Theme, Typography)
      try {
        const s = await safeStorage.get(`app_settings_${authUser.id}`);
        if (s && s.value) {
          const parsed = JSON.parse(s.value);
          setSalary(parsed.salary || 50000);
          setSalaryInput(String(parsed.salary || 50000));
          if (parsed.themeMode) setThemeMode(parsed.themeMode);
          if (parsed.fontFamily) setFontFamily(parsed.fontFamily);
          if (parsed.fontSize) setFontSize(parsed.fontSize);
          if (parsed.layoutDensity) setLayoutDensity(parsed.layoutDensity);
        }
      } catch (e) {}

      // 4. Fetch Room Membership (non-blocking — personal mode works without this)
      try {
        const { data: memberRecords, error: memberError } = await supabase
          .from('room_members')
          .select('*, rooms(*)')
          .eq('user_id', authUser.id);

        if (memberError) throw memberError;

        const activeRecord = memberRecords && memberRecords.find(r => r.status === 'accepted');
        const pendingRecord = memberRecords && memberRecords.find(r => r.status === 'pending');

        if (activeRecord) {
          setCurrentRoomId(activeRecord.room_id);
          setRoomMembershipStatus('accepted');
          if (activeRecord.rooms) {
            setCurrentRoomCode(activeRecord.rooms.invite_code);
            await fetchRoommatesAndTransactions(authUser.id, activeRecord.room_id);
          }
        } else if (pendingRecord) {
          setCurrentRoomId(pendingRecord.room_id);
          setRoomMembershipStatus('pending');
          setCurrentRoomCode(pendingRecord.rooms?.invite_code || '');
          await fetchUserPersonalTransactions(authUser.id);
        } else {
          setCurrentRoomId(null);
          setRoomMembershipStatus('none');
          setCurrentRoomCode('');
          setPendingInvites([]);
          setPendingMembers([]);
          await fetchUserPersonalTransactions(authUser.id);
        }
      } catch (e) {
        console.warn('Room membership fetch failed, showing personal ledger only', e);
        setCurrentRoomId(null);
        setRoomMembershipStatus('none');
        setPendingMembers([]);
        try { await fetchUserPersonalTransactions(authUser.id); } catch (_) {}
      }
    } catch (e) {
      console.error('loadAllSupabase error:', e);
    } finally {
      setLoading(false);
    }
  }

  async function fetchUserPersonalTransactions(userId) {
    try {
      console.log('fetchUserPersonalTransactions query initiated for userId:', userId);
      const { data, error } = await supabase
        .from('transactions')
        .select('*')
        .eq('user_id', userId)
        .eq('is_shared', false)
        .order('date', { ascending: false });
      
      if (error) {
        console.error('fetchUserPersonalTransactions query error details:', error);
        throw error;
      }
      if (data) {
        console.log('fetchUserPersonalTransactions success, records fetched:', data.length);
        const mapped = data.map(tx => {
          const match = tx.note?.match(/\[source:(\w+)\]/);
          return {
            ...tx,
            source: match ? match[1] : 'manual',
            note: tx.note ? tx.note.replace(/\[source:\w+\]/, '').trim() : ''
          };
        });
        setTransactions(mapped);
      }
    } catch (e) {
      console.error('fetchUserPersonalTransactions exception caught:', e);
    }
  }

  async function fetchRoommatesAndTransactions(userId, roomId) {
    // 1. Fetch Roommates list
    const { data: members } = await supabase
      .from('room_members')
      .select('*')
      .eq('room_id', roomId)
      .eq('status', 'accepted');
    
    if (members && members.length > 0) {
      const userIds = members.map(m => m.user_id);
      const { data: profiles } = await supabase
        .from('profiles')
        .select('*')
        .in('id', userIds);
      
      const profileMap = {};
      if (profiles) {
        profiles.forEach(p => {
          profileMap[p.id] = p;
        });
      }

      const roommatesList = members
        .filter(m => m.user_id !== userId)
        .map(m => {
          const prof = profileMap[m.user_id];
          return {
            id: m.user_id,
            name: prof?.name || 'Roommate',
            email: prof?.email || ''
          };
        });
      setRoommates(roommatesList);
    } else {
      setRoommates([]);
    }

    // 1b. Fetch Pending roommate join requests
    try {
      const { data: pendingMembersData } = await supabase
        .from('room_members')
        .select('*')
        .eq('room_id', roomId)
        .eq('status', 'pending');
      
      if (pendingMembersData && pendingMembersData.length > 0) {
        const pendingUserIds = pendingMembersData.map(m => m.user_id);
        const { data: pendingProfiles } = await supabase
          .from('profiles')
          .select('*')
          .in('id', pendingUserIds);
        
        const pendingProfileMap = {};
        if (pendingProfiles) {
          pendingProfiles.forEach(p => {
            pendingProfileMap[p.id] = p;
          });
        }

        const pendingList = pendingMembersData.map(m => {
          const prof = pendingProfileMap[m.user_id];
          return {
            id: m.user_id,
            name: prof?.name || 'Roommate',
            email: prof?.email || ''
          };
        });
        setPendingMembers(pendingList);
      } else {
        setPendingMembers([]);
      }
    } catch (e) {
      console.error("Error fetching pending members:", e);
      setPendingMembers([]);
    }

    // Load flat rent configuration
    try {
      const val = await safeStorage.get(`flat_rent_${roomId}`);
      if (val && val.value) {
        setFlatRentInput(val.value);
      } else {
        setFlatRentInput('12000');
      }
    } catch (e) {}

    // 2. Fetch Transactions (Personal + Shared room transactions)
    await fetchTransactions(userId, roomId);
  }

  async function fetchTransactions(userId, roomId) {
    try {
      let query = supabase.from('transactions').select('*').order('date', { ascending: false });

      if (roomId) {
        // Has a room: fetch personal txs + shared room txs
        query = query.or(`user_id.eq.${userId},and(room_id.eq.${roomId},is_shared.eq.true)`);
      } else {
        // No room yet: fetch only this user's personal transactions
        query = query.eq('user_id', userId).eq('is_shared', false);
      }

      const { data, error } = await query;
      if (error) throw error;
      if (data) {
        const mapped = data.map(tx => {
          const match = tx.note?.match(/\[source:(\w+)\]/);
          return {
            ...tx,
            source: match ? match[1] : 'manual',
            note: tx.note ? tx.note.replace(/\[source:\w+\]/, '').trim() : ''
          };
        });
        setTransactions(mapped);
      }
    } catch (e) {
      console.error('fetchTransactions error:', e);
    }
  }

  // --- Flat Room Actions ---
  
  async function handleCreateRoom() {
    if (!session) return;
    setRoomLoading(true);
    setErrorMsg(null);
    try {
      const inviteCode = 'FL-' + Math.floor(100000 + Math.random() * 900000);
      
      // 1. Insert room record
      const { data: room, error: rErr } = await supabase
        .from('rooms')
        .insert({ invite_code: inviteCode })
        .select()
        .single();
      
      if (rErr) throw rErr;

      // 2. Add creator as accepted member
      const { error: mErr } = await supabase
        .from('room_members')
        .insert({
          room_id: room.id,
          user_id: session.user.id,
          status: 'accepted'
        });
      
      if (mErr) throw mErr;

      setCurrentRoomId(room.id);
      setCurrentRoomCode(inviteCode);
      setRoomMembershipStatus('accepted');
      showToast('success', `Flat Created! Invite Code: ${inviteCode}`);
      await fetchRoommatesAndTransactions(session.user.id, room.id);
    } catch (e) {
      setErrorMsg(e.message || 'Could not create flat.');
    } finally {
      setRoomLoading(false);
    }
  }

  async function handleJoinRoom() {
    if (!session || !inviteCodeInput) return;
    setRoomLoading(true);
    setErrorMsg(null);
    try {
      // 1. Verify code exists
      const { data: room, error: rErr } = await supabase
        .from('rooms')
        .select('*')
        .eq('invite_code', inviteCodeInput.trim())
        .single();
      
      if (rErr || !room) {
        throw new Error('Flat Invite Code not found.');
      }

      // 2. Submit join request
      const { error: mErr } = await supabase
        .from('room_members')
        .insert({
          room_id: room.id,
          user_id: session.user.id,
          status: 'pending'
        });
      
      if (mErr) throw mErr;

      setCurrentRoomId(room.id);
      setCurrentRoomCode(room.invite_code);
      setRoomMembershipStatus('pending');
      showToast('success', 'Join request sent. Waiting for flatmate approval.');
    } catch (e) {
      setErrorMsg(e.message || 'Could not join flat.');
    } finally {
      setRoomLoading(false);
    }
  }

  async function handleAcceptInvite(roomId) {
    if (!session) return;
    setRoomLoading(true);
    try {
      const { error } = await supabase
        .from('room_members')
        .update({ status: 'accepted' })
        .match({ room_id: roomId, user_id: session.user.id });
      
      if (error) throw error;
      
      setCurrentRoomId(roomId);
      setRoomMembershipStatus('accepted');
      setPendingInvites([]);
      showToast('success', 'Invitation accepted!');
      
      // Fetch details
      const { data: room } = await supabase.from('rooms').select('invite_code').eq('id', roomId).single();
      if (room) setCurrentRoomCode(room.invite_code);
      await fetchRoommatesAndTransactions(session.user.id, roomId);
    } catch (e) {
      showToast('error', 'Error accepting invite.');
    } finally {
      setRoomLoading(false);
    }
  }

  async function handleRejectInvite(roomId) {
    if (!session) return;
    try {
      await supabase
        .from('room_members')
        .delete()
        .match({ room_id: roomId, user_id: session.user.id });
      
      setPendingInvites(prev => prev.filter(p => p.room_id !== roomId));
      showToast('success', 'Invitation declined.');
    } catch (e) {}
  }

  async function checkMembershipStatus() {
    if (!session || !currentRoomId) return;
    setRoomLoading(true);
    try {
      const { data, error } = await supabase
        .from('room_members')
        .select('*')
        .eq('room_id', currentRoomId)
        .eq('user_id', session.user.id)
        .single();
      
      if (data) {
        setRoomMembershipStatus(data.status);
        if (data.status === 'accepted') {
          showToast('success', 'Membership approved!');
          await fetchRoommatesAndTransactions(session.user.id, currentRoomId);
        } else {
          showToast('error', 'Still pending roommate approval.');
        }
      }
    } catch (e) {}
    setRoomLoading(false);
  }

  async function handleLeaveRoom() {
    if (!session || !currentRoomId) return;
    if (!window.confirm("Leave this roommate flat? You will lose access to shared splits.")) return;
    try {
      await supabase
        .from('room_members')
        .delete()
        .match({ room_id: currentRoomId, user_id: session.user.id });
      
      setCurrentRoomId(null);
      setCurrentRoomCode('');
      setRoommates([]);
      setPendingMembers([]);
      setTransactions([]);
      showToast('success', 'Left roommate flat.');
      await fetchUserPersonalTransactions(session.user.id);
    } catch (e) {}
  }

  // Realtime subscription hook
  useEffect(() => {
    if (!session || !currentRoomId || roomMembershipStatus !== 'accepted') return;

    const channel = supabase
      .channel('realtime-transactions')
      .on('postgres_changes', { 
        event: '*', 
        schema: 'public', 
        table: 'transactions',
        filter: `room_id=eq.${currentRoomId}`
      }, () => {
        fetchTransactions(session.user.id, currentRoomId);
      })
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, [session, currentRoomId, roomMembershipStatus]);

  // --- Seed Demo Data (Saves to Supabase Database) ---
  const seedDemoData = async () => {
    if (!session || !currentRoomId) return;
    const today = new Date();
    const demoTxs = [];
    
    for (let i = 45; i >= 0; i--) {
      const d = new Date(today);
      d.setDate(today.getDate() - i);
      const dateStr = isoDate(d);
      const dayOfWeek = d.getDay();

      // Monthly Rent split (Shared roommates expense - logged on 1st of month by Siva/User)
      if (d.getDate() === 1) {
        demoTxs.push({
          user_id: session.user.id,
          room_id: currentRoomId,
          category: 'Rent',
          amount: 18000, 
          merchant: 'Flat Owner',
          note: 'House rent split equally [source:manual]',
          is_shared: true, 
          logged_by: currentUser.name,
          date: dateStr
        });
      }

      // personal Tea/Snacks (Siva leakage source)
      demoTxs.push({
        user_id: session.user.id,
        room_id: null, // Personal individual expense
        category: 'Food',
        amount: 60,
        merchant: 'Tapri Chai & Samosa',
        note: 'Daily office evening tea [source:manual]',
        is_shared: false,
        logged_by: currentUser.name,
        date: dateStr
      });

      // Commute
      if (dayOfWeek >= 1 && dayOfWeek <= 5) {
        demoTxs.push({
          user_id: session.user.id,
          room_id: null,
          category: 'Transport',
          amount: 100 + Math.floor(Math.random() * 50),
          merchant: 'Uber Auto',
          note: 'Commute to office [source:manual]',
          is_shared: false,
          logged_by: currentUser.name,
          date: dateStr
        });
      }
    }

    try {
      const { error } = await supabase.from('transactions').insert(demoTxs);
      if (error) {
        console.error('seedDemoData Supabase insert error details:', error);
        throw error;
      }
      showToast('success', 'Seeded demo transactions to Supabase!');
      await fetchTransactions(session.user.id, currentRoomId);
    } catch (e) {
      showToast('error', 'Could not seed demo data.');
    }
  };

  const clearAllData = async () => {
    if (!session) return;
    if (!window.confirm('Delete all your ledger entries from Supabase? This cannot be undone.')) return;
    
    try {
      const { error } = await supabase
        .from('transactions')
        .delete()
        .eq('user_id', session.user.id);
      
      if (error) throw error;
      
      setTransactions([]);
      showToast('success', 'Ledger entries deleted.');
    } catch (e) {
      showToast('error', 'Could not clear data.');
    }
  };

  // --- Roommates Manager (Invite by email) ---
  const handleAddRoommate = async (e) => {
    e.preventDefault();
    if (!authForm.email || !currentRoomId) {
      showToast('error', 'Please enter roommate email.');
      return;
    }
    
    setRoomLoading(true);
    try {
      // 1. Resolve email to profile User ID
      const { data: profile, error } = await supabase
        .from('profiles')
        .select('id, name')
        .eq('email', authForm.email.trim().toLowerCase())
        .single();
      
      if (error || !profile) {
        throw new Error('Roommate account not found. Ask them to sign up first!');
      }

      // 2. Insert pending roommate membership request
      const { error: inviteErr } = await supabase
        .from('room_members')
        .insert({
          room_id: currentRoomId,
          user_id: profile.id,
          status: 'pending'
        });
      
      if (inviteErr) {
        throw new Error('Roommate invitation is already pending or they are already a member.');
      }

      showToast('success', `Sent pending flat invitation to ${profile.name}!`);
      setIsAuthOpen(false);
      setAuthForm({ name: '', email: '' });
    } catch (e) {
      showToast('error', e.message || 'Error inviting roommate.');
    } finally {
      setRoomLoading(false);
    }
  };

  // Approve a pending roommate request from the host dashboard
  const handleApproveRoommateRequest = async (roommateUserId) => {
    if (!currentRoomId) return;
    try {
      const { error } = await supabase
        .from('room_members')
        .update({ status: 'accepted' })
        .match({ room_id: currentRoomId, user_id: roommateUserId });
      
      if (error) throw error;
      showToast('success', 'Roommate request approved!');
      await fetchRoommatesAndTransactions(session.user.id, currentRoomId);
    } catch (e) {
      showToast('error', 'Error approving request.');
    }
  };

  const saveSalary = async () => {
    if (!session) return;
    const val = parseFloat(salaryInput) || 0;
    setSalary(val);
    try {
      await safeStorage.set(`app_settings_${session.user.id}`, JSON.stringify({ 
        salary: val, themeMode, fontFamily, fontSize, layoutDensity 
      }));
      showToast('success', `Monthly personal income set to ${fmt(val)}`);
    } catch (e) {
      showToast('error', 'Could not save income.');
    }
  };

  const saveFlatRent = async () => {
    if (!session || !currentRoomId) return;
    if (roomAdminId && session.user.id !== roomAdminId) {
      showToast('error', 'Only the Group Admin can configure flat rent.');
      return;
    }
    const val = parseFloat(flatRentInput) || 0;
    try {
      // 1. Delete existing AdminConfig transaction for this room
      const { error: deleteErr } = await supabase
        .from('transactions')
        .delete()
        .eq('room_id', currentRoomId)
        .eq('category', 'System')
        .eq('merchant', 'AdminConfig');
      if (deleteErr) throw deleteErr;

      // 2. Insert new AdminConfig transaction
      const targetAdminId = roomAdminId || session.user.id;
      const tx = {
        user_id: targetAdminId,
        room_id: currentRoomId,
        category: 'System',
        amount: val,
        merchant: 'AdminConfig',
        note: `admin_id:${targetAdminId}`,
        is_shared: true,
        logged_by: 'System',
        date: '2000-01-01'
      };
      const { error: insertErr } = await supabase.from('transactions').insert(tx);
      if (insertErr) throw insertErr;

      // 3. Clear any manually logged Rent transactions for the current month in this room
      // to ensure the new configured amount is clean and takes effect immediately.
      // We do this by extracting matching IDs from local transactions state to run a clean,
      // robust primary-key delete, avoiding PostgREST DELETE errors with like filters.
      const curMonth = isoDate(new Date()).slice(0, 7);
      const rentIdsToDelete = transactions
        .filter(t => t.room_id === currentRoomId && t.category === 'Rent' && t.is_shared && t.date.slice(0, 7) === curMonth)
        .map(t => t.id);

      if (rentIdsToDelete.length > 0) {
        const { error: clearErr } = await supabase
          .from('transactions')
          .delete()
          .in('id', rentIdsToDelete);
        if (clearErr) throw clearErr;
      }

      await safeStorage.set(`flat_rent_${currentRoomId}`, String(val));
      showToast('success', `Monthly Flat Rent configured to ₹${val}`);
      await fetchTransactions(session.user.id, currentRoomId);
    } catch (e) {
      showToast('error', 'Could not save flat rent.');
    }
  };

  const quickLogRent = async () => {
    if (!session || !currentRoomId) return;
    if (roomAdminId && session.user.id !== roomAdminId) {
      showToast('error', 'Only the Group Admin can log rent splits.');
      return;
    }
    const amt = parseFloat(flatRentInput) || 12000;
    const dateStr = isoDate(new Date());
    
    try {
      // 1. Delete any existing manually logged Rent transactions for the current month in this room
      // to prevent duplicate logging.
      // We extract matching IDs from local transactions state to run a clean,
      // robust primary-key delete, avoiding PostgREST DELETE errors with like filters.
      const curMonth = dateStr.slice(0, 7);
      const rentIdsToDelete = transactions
        .filter(t => t.room_id === currentRoomId && t.category === 'Rent' && t.is_shared && t.date.slice(0, 7) === curMonth)
        .map(t => t.id);

      if (rentIdsToDelete.length > 0) {
        const { error: deleteErr } = await supabase
          .from('transactions')
          .delete()
          .in('id', rentIdsToDelete);
        if (deleteErr) throw deleteErr;
      }

      // 2. Insert the new rent transaction
      const tx = {
        user_id: session.user.id,
        room_id: currentRoomId,
        category: 'Rent',
        amount: amt,
        merchant: 'Flat Owner',
        note: 'House rent split equally [source:manual]',
        is_shared: true,
        logged_by: currentUser?.name || 'Host',
        date: dateStr
      };

      const { error: insertErr } = await supabase.from('transactions').insert(tx);
      if (insertErr) {
        console.error('quickLogRent Supabase error details:', insertErr);
        throw insertErr;
      }
      showToast('success', `Logged monthly rent transaction of ₹${amt}!`);
      await fetchTransactions(session.user.id, currentRoomId);
    } catch (e) {
      showToast('error', 'Error logging rent transaction.');
    }
  };

  const transferAdminRole = async (newAdminId, newAdminName) => {
    if (!session || !currentRoomId || !adminConfigTx) return;
    if (session.user.id !== roomAdminId) {
      showToast('error', 'Only the current Group Admin can transfer administrative role.');
      return;
    }
    
    try {
      // 1. Delete existing AdminConfig transaction for this room
      const { error: deleteErr } = await supabase
        .from('transactions')
        .delete()
        .eq('room_id', currentRoomId)
        .eq('category', 'System')
        .eq('merchant', 'AdminConfig');
      if (deleteErr) throw deleteErr;

      // 2. Insert new AdminConfig transaction with newAdminId
      const tx = {
        user_id: newAdminId,
        room_id: currentRoomId,
        category: 'System',
        amount: adminConfigTx.amount,
        merchant: 'AdminConfig',
        note: `admin_id:${newAdminId}`,
        is_shared: true,
        logged_by: 'System',
        date: '2000-01-01'
      };
      const { error: insertErr } = await supabase.from('transactions').insert(tx);
      if (insertErr) throw insertErr;

      showToast('success', `Admin role successfully transferred to ${newAdminName}!`);
      await fetchTransactions(session.user.id, currentRoomId);
    } catch (e) {
      showToast('error', 'Could not transfer admin role. Please try again.');
    }
  };

  // Sync flatRentInput from database config when loaded
  useEffect(() => {
    if (adminConfigTx) {
      setFlatRentInput(String(adminConfigTx.amount));
    }
  }, [adminConfigTx]);

  // Auto-upgrade missing config
  useEffect(() => {
    if (!session || !currentRoomId || !transactions || loading) return;
    
    const configExists = transactions.some(t => 
      t.is_shared && 
      t.category === 'System' && 
      t.merchant === 'AdminConfig'
    );

    if (!configExists) {
      const initAdminConfig = async () => {
        const { data: existing } = await supabase
          .from('transactions')
          .select('id')
          .eq('room_id', currentRoomId)
          .eq('category', 'System')
          .eq('merchant', 'AdminConfig')
          .limit(1);
        
        if (existing && existing.length > 0) return;

        const { data: members } = await supabase
          .from('room_members')
          .select('user_id')
          .eq('room_id', currentRoomId)
          .eq('status', 'accepted');
        
        let adminId = session.user.id;
        if (members && members.length > 0) {
          const hasUs = members.some(m => m.user_id === session.user.id);
          if (!hasUs) {
            adminId = members[0].user_id;
          }
        }

        const initialRent = '12000';
        const tx = {
          user_id: adminId,
          room_id: currentRoomId,
          category: 'System',
          amount: parseFloat(initialRent),
          merchant: 'AdminConfig',
          note: `admin_id:${adminId}`,
          is_shared: true,
          logged_by: 'System',
          date: '2000-01-01'
        };

        try {
          await supabase.from('transactions').insert(tx);
          await fetchTransactions(session.user.id, currentRoomId);
        } catch (e) {
          console.error("Failed to auto-create AdminConfig:", e);
        }
      };
      
      initAdminConfig();
    }
  }, [currentRoomId, transactions, session, loading]);

  // --- Transactions Logging ---
  const addTransaction = async (e) => {
    if (e) e.preventDefault();
    if (!session) return;

    const amt = parseFloat(form.amount);
    if (!amt || amt <= 0) {
      showToast('error', 'Please enter a valid amount.');
      return;
    }
    if (!form.merchant) {
      showToast('error', 'Please specify a merchant.');
      return;
    }

    // Guard: shared transaction requires a flat (room) to exist
    if (form.isShared && !currentRoomId) {
      showToast('error', 'You must create or join a flat before adding shared expenses.');
      return;
    }

    // Append source metadata to note
    const noteText = form.note.trim();
    const noteWithSource = noteText ? `${noteText} [source:${form.source || 'manual'}]` : `[source:${form.source || 'manual'}]`;

    const newTx = {
      user_id: session.user.id,
      ...(form.isShared && { room_id: currentRoomId }),
      category: form.category,
      amount: amt,
      merchant: form.merchant.trim(),
      note: noteWithSource,
      is_shared: form.isShared,
      logged_by: currentUser?.name || 'Roommate',
      date: form.date
    };

    try {
      const { error } = await supabase.from('transactions').insert(newTx);
      if (error) {
        console.error('addTransaction Supabase insert error details:', error);
        throw error;
      }

      showToast('success', `Logged ${fmt(amt)} spent at ${newTx.merchant}`);
      await fetchTransactions(session.user.id, currentRoomId);

      setForm({
        ...form,
        amount: '',
        merchant: '',
        note: '',
        source: 'manual',
        isShared: false
      });
    } catch (e) {
      showToast('error', 'Could not log transaction.');
    }
  };

  const deleteTransaction = async (id) => {
    if (!session) return;
    try {
      const { error } = await supabase.from('transactions').delete().eq('id', id);
      if (error) throw error;
      showToast('success', 'Entry deleted.');
      await fetchTransactions(session.user.id, currentRoomId);
    } catch (e) {
      showToast('error', 'Could not delete entry.');
    }
  };

  // --- Dates Variables ---
  const todayStr = isoDate(new Date());
  const curMonthStr = todayStr.slice(0, 7);
  const curYearStr = todayStr.slice(0, 4);

  // --- Calculation Engines ---
  const personalTransactions = useMemo(() => {
    return transactions
      .filter(tx => !tx.is_shared && tx.user_id === session?.user?.id && tx.category !== 'System')
      .sort((a, b) => b.date.localeCompare(a.date));
  }, [transactions, session]);

  const sharedTransactions = useMemo(() => {
    return transactions.filter(t => t.is_shared && t.category !== 'System').sort((a,b) => b.date.localeCompare(a.date));
  }, [transactions]);

  const activeAnalysisTxs = useMemo(() => {
    return analysisType === 'personal' ? personalTransactions : sharedTransactions;
  }, [analysisType, personalTransactions, sharedTransactions]);

  const processedTxs = useMemo(() => {
    return activeAnalysisTxs.filter(t => {
      const matchCat = filterCategory === 'All' || t.category === filterCategory;
      const matchSource = filterSource === 'All' || t.source === filterSource;
      const matchSearch = searchQuery === '' || 
        t.merchant.toLowerCase().includes(searchQuery.toLowerCase()) || 
        (t.note && t.note.toLowerCase().includes(searchQuery.toLowerCase())) ||
        t.category.toLowerCase().includes(searchQuery.toLowerCase());
      return matchCat && matchSource && matchSearch;
    });
  }, [activeAnalysisTxs, filterCategory, filterSource, searchQuery]);



  const totals = useMemo(() => {
    let t = 0, w = 0, m = 0, y = 0;
    const todayObj = new Date();
    const day = (todayObj.getDay() + 6) % 7;
    const startOfWeekObj = new Date(todayObj);
    startOfWeekObj.setDate(todayObj.getDate() - day);
    startOfWeekObj.setHours(0,0,0,0);
    const startOfWeekStr = isoDate(startOfWeekObj);

    for (const tx of activeAnalysisTxs) {
      if (tx.date === todayStr) t += tx.amount;
      if (tx.date >= startOfWeekStr) w += tx.amount;
      if (tx.date.slice(0, 7) === curMonthStr) m += tx.amount;
      if (tx.date.slice(0, 4) === curYearStr) y += tx.amount;
    }

    if (analysisType === 'roommates' && shouldInjectRent) {
      m += activeRentAmount;
      y += activeRentAmount;
    }

    return { today: t, week: w, month: m, year: y };
  }, [activeAnalysisTxs, todayStr, curMonthStr, curYearStr, analysisType, shouldInjectRent, activeRentAmount]);

  const dailyBudget = salary / 30;
  const overToday = totals.today - dailyBudget;
  const remainingSalary = salary - totals.month;

  const categoryBreakdown = useMemo(() => {
    const map = {};
    const monthlyTxs = activeAnalysisTxs.filter(t => t.date.slice(0, 7) === curMonthStr);
    for (const tx of monthlyTxs) {
      const amt = (analysisType === 'roommates' && tx.is_shared) ? (tx.amount / (roommates.length + 1)) : tx.amount;
      map[tx.category] = (map[tx.category] || 0) + amt;
    }

    if (analysisType === 'roommates' && shouldInjectRent) {
      const userRentShare = activeRentAmount / (roommates.length + 1);
      map['Rent'] = (map['Rent'] || 0) + userRentShare;
    }

    return Object.entries(map)
      .map(([name, amount]) => ({ name, amount, color: CATEGORY_COLORS[name] || '#4B5563' }))
      .sort((a, b) => b.amount - a.amount);
  }, [activeAnalysisTxs, curMonthStr, roommates, analysisType, shouldInjectRent, activeRentAmount]);

  const totalMonthlySpend = useMemo(() => {
    return categoryBreakdown.reduce((sum, c) => sum + c.amount, 0);
  }, [categoryBreakdown]);

  const doughnutSegments = useMemo(() => {
    if (totalMonthlySpend === 0) return [];
    let accumulatedPercent = 0;
    return categoryBreakdown.map(c => {
      const percent = (c.amount / totalMonthlySpend) * 100;
      const startPercent = accumulatedPercent;
      accumulatedPercent += percent;
      
      const r = 50;
      const circ = 2 * Math.PI * r;
      const strokeDasharray = `${(percent / 100) * circ} ${circ}`;
      const strokeDashoffset = `${circ - (startPercent / 100) * circ}`;

      return {
        ...c,
        percent,
        strokeDasharray,
        strokeDashoffset
      };
    });
  }, [categoryBreakdown, totalMonthlySpend]);

  // --- Roommate Splits / Who Owes Who Matrix ---
  const roommateDues = useMemo(() => {
    const N = roommates.length + 1;
    const todayObj = new Date();
    const day = (todayObj.getDay() + 6) % 7;
    const startOfWeekObj = new Date(todayObj);
    startOfWeekObj.setDate(todayObj.getDate() - day);
    startOfWeekObj.setHours(0,0,0,0);
    const startOfWeekStr = isoDate(startOfWeekObj);

    // Initialize member lists and stats
    const memberNames = [];
    if (currentUser) memberNames.push(currentUser.name);
    roommates.forEach(r => memberNames.push(r.name));

    const totalPaidMap = {};
    const txCountMap = {};
    memberNames.forEach(name => {
      totalPaidMap[name] = 0;
      txCountMap[name] = 0;
    });

    let totalRoomExpense = 0;
    let monthlyRoomExpense = 0;
    let weeklyRoomExpense = 0;
    let dailyRoomExpense = 0;

    for (const tx of transactions) {
      if (tx.is_shared && tx.category !== 'System') {
        totalRoomExpense += tx.amount;
        if (tx.date === todayStr) dailyRoomExpense += tx.amount;
        if (tx.date >= startOfWeekStr) weeklyRoomExpense += tx.amount;
        if (tx.date.slice(0, 7) === curMonthStr) monthlyRoomExpense += tx.amount;

        const payer = tx.logged_by;
        if (totalPaidMap[payer] !== undefined) {
          totalPaidMap[payer] += tx.amount;
          txCountMap[payer] = (txCountMap[payer] || 0) + 1;
        }
      }
    }

    if (shouldInjectRent) {
      totalRoomExpense += activeRentAmount;
      monthlyRoomExpense += activeRentAmount;
      // Do not credit any member (including the admin) for the auto-injected rent
      // so it remains un-attributed and excluded from peer-to-peer settlements.
    }

    const otherSharedTotal = totalRoomExpense - (shouldInjectRent ? activeRentAmount : 0);
    const otherShare = N > 0 ? otherSharedTotal / N : 0;
    const rentShare = shouldInjectRent ? (activeRentAmount / N) : 0;
    const totalShare = otherShare + rentShare;

    const balances = {};
    const memberSummaries = memberNames.map(name => {
      const paid = totalPaidMap[name] || 0;
      // The settlement balance covers only other shared expenses (groceries, bills, etc.)
      const bal = paid - otherShare;
      balances[name] = Math.round(bal);

      let status = 'Settled';
      if (bal > 0) status = 'Owed';
      else if (bal < 0) status = 'Owes';

      return {
        name,
        paid: Math.round(paid),
        share: Math.round(totalShare), // Display total share including rent split
        balance: Math.round(bal),
        amountToPay: bal < 0 ? Math.round(-bal) : 0,
        amountToReceive: bal > 0 ? Math.round(bal) : 0,
        status
      };
    });

    // Debt settlements simplified path:
    const debtors = [];
    const creditors = [];
    Object.entries(balances).forEach(([name, bal]) => {
      if (bal < 0) debtors.push({ name, bal: -bal });
      if (bal > 0) creditors.push({ name, bal });
    });

    const duesList = [];
    let dIdx = 0;
    let cIdx = 0;
    while (dIdx < debtors.length && cIdx < creditors.length) {
      const debtor = debtors[dIdx];
      const creditor = creditors[cIdx];
      const transfer = Math.min(debtor.bal, creditor.bal);
      duesList.push({
        from: debtor.name,
        to: creditor.name,
        amount: Math.round(transfer)
      });
      debtor.bal -= transfer;
      creditor.bal -= transfer;
      if (debtor.bal <= 1) dIdx++;
      if (creditor.bal <= 1) cIdx++;
    }

    // Determine stats
    let highestContributor = { name: 'N/A', amount: 0 };
    let lowestContributor = { name: 'N/A', amount: Infinity };
    let mostActiveBuyer = { name: 'N/A', count: 0 };

    if (memberNames.length > 0) {
      let maxPaid = -1;
      let minPaid = Infinity;
      let maxCount = -1;

      memberNames.forEach(name => {
        const paid = totalPaidMap[name] || 0;
        const count = txCountMap[name] || 0;

        if (paid > maxPaid) {
          maxPaid = paid;
          highestContributor = { name, amount: Math.round(paid) };
        }
        if (paid < minPaid) {
          minPaid = paid;
          lowestContributor = { name, amount: Math.round(paid) };
        }
        if (count > maxCount) {
          maxCount = count;
          mostActiveBuyer = { name, count };
        }
      });
    }

    if (lowestContributor.amount === Infinity) {
      lowestContributor = { name: 'N/A', amount: 0 };
    }

    return {
      balances,
      duesList,
      totalRoomExpense,
      monthlyRoomExpense,
      weeklyRoomExpense,
      dailyRoomExpense,
      memberSummaries,
      highestContributor,
      lowestContributor,
      mostActiveBuyer
    };
  }, [transactions, roommates, currentUser, todayStr, curMonthStr, shouldInjectRent, activeRentAmount, roomAdminName]);

  const healthScore = useMemo(() => {
    if (salary === 0) return 0;
    
    const savings = salary - totalMonthlySpend;
    const savingsRate = (savings / salary) * 100;
    let savingsScore = 0;
    if (savingsRate >= 35) savingsScore = 100;
    else if (savingsRate > 0) savingsScore = (savingsRate / 35) * 100;

    const monthlyTxs = personalTransactions.filter(t => t.date.slice(0, 7) === curMonthStr);
    const fixedBills = monthlyTxs.filter(t => t.category === 'Rent' || t.category === 'Bills').reduce((sum, t) => sum + t.amount, 0);
    const fixedRatio = (fixedBills / salary) * 100;
    let fixedScore = 0;
    if (fixedRatio <= 25) fixedScore = 100;
    else if (fixedRatio < 50) fixedScore = 100 - ((fixedRatio - 25) / 25) * 100;

    const totalInvestments = personalTransactions.filter(t => t.category === 'Investments').reduce((sum, t) => sum + t.amount, 0);
    const liquidAssets = totalInvestments + remainingSalary;
    const avgExpense = salary * 0.65;
    const targetBuffer = avgExpense * 4;
    const emergencyScore = Math.min((liquidAssets / targetBuffer) * 100, 100);

    const discSpent = monthlyTxs
      .filter(t => ['Food', 'Shopping', 'Entertainment'].includes(t.category))
      .reduce((sum, t) => sum + t.amount, 0);
    const discRatio = (discSpent / salary) * 100;
    let discScore = 0;
    if (discRatio <= 25) discScore = 100;
    else if (discRatio < 60) discScore = 100 - ((discRatio - 25) / 35) * 100;

    const finalScore = (savingsScore * 0.35) + (fixedScore * 0.2) + (emergencyScore * 0.2) + (discScore * 0.25);
    return Math.round(Math.max(0, Math.min(finalScore, 100)));
  }, [salary, totalMonthlySpend, personalTransactions, curMonthStr, remainingSalary]);

  const riskStatus = useMemo(() => {
    if (healthScore >= 80) return { label: 'LOW RISK', color: 'var(--positive)', desc: 'Your bachelor budget is robust. High savings rate and liquid emergency funds.' };
    if (healthScore >= 55) return { label: 'MEDIUM RISK', color: '#D97706', desc: 'Decent control, but watch out for lifestyle creep (frequent restaurant orders/online sales).' };
    if (healthScore >= 35) return { label: 'HIGH RISK', color: 'var(--stamp)', desc: 'Over 50% of salary goes to discretionary dine-outs and gadgets. Room rent and bills take up significant margin.' };
    return { label: 'CRITICAL RISK', color: 'var(--stamp)', desc: 'Paycheck to paycheck. Discretionary spending exceeds total savings capacity. Minimize food orders immediately.' };
  }, [healthScore]);

  const survivalData = useMemo(() => {
    const today = new Date();
    const dayOfMonth = today.getDate();
    const totalDays = getDaysInMonth(today.getFullYear(), today.getMonth());
    
    const burnRatePerDay = dayOfMonth > 0 ? totalMonthlySpend / dayOfMonth : 0;
    const projectedTotalSpend = burnRatePerDay * totalDays;
    const remainingDaysPace = burnRatePerDay > 0 ? (salary - totalMonthlySpend) / burnRatePerDay : 99;
    
    const isOverbudget = projectedTotalSpend > salary;
    const runoutDay = Math.min(totalDays, Math.max(1, Math.round(totalMonthlySpend / (burnRatePerDay || 1))));

    return {
      dayOfMonth,
      totalDays,
      burnRatePerDay: Math.round(burnRatePerDay),
      projectedTotalSpend: Math.round(projectedTotalSpend),
      isOverbudget,
      remainingDaysPace: Math.max(0, Math.round(remainingDaysPace)),
      runoutDay: isOverbudget ? Math.round(salary / (burnRatePerDay || 1)) : totalDays
    };
  }, [totalMonthlySpend, salary]);

  const leakageItems = useMemo(() => {
    const groups = {};
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    const thirtyDaysAgoStr = isoDate(thirtyDaysAgo);

    const recentSmallTxs = personalTransactions.filter(t => t.date >= thirtyDaysAgoStr && t.amount <= 250);

    for (const tx of recentSmallTxs) {
      const key = `${tx.category}-${tx.merchant.toLowerCase()}`;
      if (!groups[key]) {
        groups[key] = { category: tx.category, merchant: tx.merchant, count: 0, total: 0, items: [] };
      }
      groups[key].count++;
      groups[key].total += tx.amount;
      groups[key].items.push(tx);
    }

    return Object.values(groups)
      .filter(g => g.count >= 5)
      .map(g => {
        const monthlyCost = g.total;
        const yearlyCost = monthlyCost * 12;
        const r = 0.12;
        const n = 5;
        const pmt = monthlyCost;
        const fv = pmt * ((Math.pow(1 + r/12, n*12) - 1) / (r/12));
        return {
          ...g,
          monthlyCost,
          yearlyCost,
          projectedWealth: Math.round(fv)
        };
      });
  }, [personalTransactions]);

  const subscriptions = useMemo(() => {
    const groups = {};
    const sortedTxs = [...personalTransactions].sort((a,b) => a.date.localeCompare(b.date));
    
    for (const tx of sortedTxs) {
      if (tx.category === 'Rent' || tx.category === 'Investments') continue;
      
      const key = `${tx.merchant.toLowerCase()}-${tx.amount}`;
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(tx);
    }

    const detected = [];
    Object.entries(groups).forEach(([key, txList]) => {
      if (txList.length >= 2) {
        let totalGap = 0;
        let gapCount = 0;
        for (let i = 1; i < txList.length; i++) {
          const d1 = new Date(txList[i-1].date);
          const d2 = new Date(txList[i].date);
          const diffDays = Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
          totalGap += diffDays;
          gapCount++;
        }
        const avgGap = gapCount > 0 ? totalGap / gapCount : 0;
        
        if (avgGap >= 25 && avgGap <= 35) {
          const sample = txList[0];
          const name = sample.merchant;
          const usage = subscriptionUsage[name] !== undefined ? subscriptionUsage[name] : 10;
          const isLowUsage = usage <= 3; 
          detected.push({
            name,
            amount: sample.amount,
            category: sample.category,
            avgGap: Math.round(avgGap),
            occurrences: txList.length,
            usage,
            isLowUsage
          });
        }
      }
    });

    return detected;
  }, [personalTransactions, subscriptionUsage]);

  const discretionarySum = useMemo(() => {
    const monthlyTxs = personalTransactions.filter(t => t.date.slice(0, 7) === curMonthStr);
    const targetCats = ['Food', 'Shopping', 'Entertainment'];
    return monthlyTxs
      .filter(t => targetCats.includes(t.category))
      .reduce((sum, t) => sum + t.amount, 0);
  }, [personalTransactions, curMonthStr]);

  const discretionarySaved = Math.round(discretionarySum * (reductionPercent / 100));
  const wealthLossProjection5Yr = useMemo(() => {
    const r = 0.12;
    const n = 5;
    const pmt = discretionarySaved;
    return Math.round(pmt * ((Math.pow(1 + r/12, n*12) - 1) / (r/12)));
  }, [discretionarySaved]);

  const wealthLossProjection10Yr = useMemo(() => {
    const r = 0.12;
    const n = 10;
    const pmt = discretionarySaved;
    return Math.round(pmt * ((Math.pow(1 + r/12, n*12) - 1) / (r/12)));
  }, [discretionarySaved]);

  // --- ML Forecast Simulation ---
  const chartAndForecastData = useMemo(() => {
    const monthlyData = {};
    const today = new Date();
    
    for (let i = 5; i >= 0; i--) {
      const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
      const mKey = d.toISOString().slice(0, 7);
      monthlyData[mKey] = { label: d.toLocaleString('default', { month: 'short' }), amount: 0, isForecast: false };
    }

    for (const tx of activeAnalysisTxs) {
      const mKey = tx.date.slice(0, 7);
      if (monthlyData[mKey]) {
        const amt = (analysisType === 'roommates' && tx.is_shared) ? (tx.amount / (roommates.length + 1)) : tx.amount;
        monthlyData[mKey].amount += amt;
      }
    }

    // In roommates mode, auto-inject rent share for months that don't have it logged
    if (analysisType === 'roommates' && activeRentAmount > 0) {
      const monthsWithRent = new Set();
      for (const tx of activeAnalysisTxs) {
        if (tx.is_shared && tx.category === 'Rent') {
          monthsWithRent.add(tx.date.slice(0, 7));
        }
      }
      
      const userRentShare = activeRentAmount / (roommates.length + 1);
      for (const mKey of Object.keys(monthlyData)) {
        if (!monthsWithRent.has(mKey)) {
          monthlyData[mKey].amount += userRentShare;
        }
      }
    }

    const dataPoints = Object.entries(monthlyData).map(([key, val]) => ({
      key,
      ...val
    })).sort((a,b) => a.key.localeCompare(b.key));

    const forecastPoints = [];
    let history = dataPoints.map(d => d.amount);
    
    let trend = 0;
    if (history.length >= 2) {
      trend = (history[history.length - 1] - history[history.length - 2]) * 0.15;
    }

    for (let i = 1; i <= 3; i++) {
      const futDate = new Date(today.getFullYear(), today.getMonth() + i, 1);
      const futKey = futDate.toISOString().slice(0, 7);
      
      const n = history.length;
      const avg = (history[n - 1] + (history[n - 2] || 0) + (history[n - 3] || 0)) / (n >= 3 ? 3 : n);
      const forecastVal = Math.max(salary * 0.4, Math.round(avg + trend * i));
      
      forecastPoints.push({
        key: futKey,
        label: futDate.toLocaleString('default', { month: 'short' }) + '*',
        amount: forecastVal,
        isForecast: true
      });
      history.push(forecastVal);
    }

    return [...dataPoints, ...forecastPoints];
  }, [activeAnalysisTxs, salary, roommates, analysisType, activeRentAmount]);

  const svgLineChartPath = useMemo(() => {
    const pts = chartAndForecastData;
    if (pts.length === 0) return null;
    const maxVal = Math.max(...pts.map(p => p.amount), salary, 10000) * 1.1;
    
    const width = 600;
    const height = 150;
    const padX = 40;
    const padY = 20;

    const coords = pts.map((p, idx) => {
      const x = padX + (idx / (pts.length - 1)) * (width - 2 * padX);
      const y = height - padY - (p.amount / maxVal) * (height - 2 * padY);
      return { x, y, isForecast: p.isForecast };
    });

    const getBezierPath = (points) => {
      if (points.length === 0) return '';
      let path = `M ${points[0].x} ${points[0].y}`;
      for (let i = 1; i < points.length; i++) {
        const p0 = points[i - 1];
        const p1 = points[i];
        const cp1x = p0.x + (p1.x - p0.x) / 3;
        const cp1y = p0.y;
        const cp2x = p0.x + 2 * (p1.x - p0.x) / 3;
        const cp2y = p1.y;
        path += ` C ${cp1x} ${cp1y}, ${cp2x} ${cp2y}, ${p1.x} ${p1.y}`;
      }
      return path;
    };

    const histCoords = coords.filter(c => !c.isForecast);
    const foreCoords = [coords[histCoords.length - 1], ...coords.filter(c => c.isForecast)];

    const histPath = getBezierPath(histCoords);
    const forePath = getBezierPath(foreCoords);

    const histAreaPath = histCoords.length > 0 
      ? `${histPath} L ${histCoords[histCoords.length - 1].x} 130 L ${histCoords[0].x} 130 Z` 
      : '';
    const foreAreaPath = foreCoords.length > 0 
      ? `${forePath} L ${foreCoords[foreCoords.length - 1].x} 130 L ${foreCoords[0].x} 130 Z` 
      : '';

    return {
      histPath,
      forePath,
      histAreaPath,
      foreAreaPath,
      coords,
      maxVal,
      width,
      height
    };
  }, [chartAndForecastData, salary]);

  const mistakeAlerts = useMemo(() => {
    const alerts = [];
    
    const months = {};
    for (const t of personalTransactions) {
      const m = t.date.slice(0, 7);
      if (!months[m]) months[m] = { total: 0, discretionary: 0 };
      months[m].total += t.amount;
      if (['Food', 'Shopping', 'Entertainment'].includes(t.category)) {
        months[m].discretionary += t.amount;
      }
    }

    const sortedMonths = Object.keys(months).sort();
    if (sortedMonths.length >= 3) {
      const m1 = months[sortedMonths[sortedMonths.length - 3]];
      const m2 = months[sortedMonths[sortedMonths.length - 2]];
      const m3 = months[sortedMonths[sortedMonths.length - 1]];
      
      const discTrend1 = m2.discretionary - m1.discretionary;
      const discTrend2 = m3.discretionary - m2.discretionary;
      
      if (discTrend1 > 0 && discTrend2 > 0) {
        alerts.push({
          type: 'lifestyle',
          title: 'Creeping Dineouts & Delivery',
          desc: 'Discretionary spend (food, shopping) has creeped up consecutively for 3 months. Cook at the flat to de-risk your saving rate.'
        });
      }
    }

    const essentialSpent = personalTransactions
      .filter(t => t.date.slice(0,7) === curMonthStr && ['Rent', 'Bills'].includes(t.category))
      .reduce((sum, t) => sum + t.amount, 0);
    if (salary > 0 && (essentialSpent / salary) > 0.40) {
      alerts.push({
        type: 'fixed_ratio',
        title: 'Commitment Overhead Over 40%',
        desc: `Shared rent split and WiFi bills consume ${((essentialSpent/salary)*100).toFixed(0)}% of your salary. Consider a cheaper flat configuration.`
      });
    }

    const subCount = subscriptions.filter(s => s.isLowUsage).length;
    if (subCount >= 2) {
      alerts.push({
        type: 'subs_bloat',
        title: 'Zombie Subscription Waste',
        desc: `You have ${subCount} subscriptions marked with very low usage (<= 3 logs/month) draining liquid cash.`
      });
    }

    return alerts;
  }, [personalTransactions, curMonthStr, salary, subscriptions]);

  const askAI = (queryText) => {
    const q = queryText.toLowerCase();
    let reply = "";

    if (q.includes('leak') || q.includes('waste') || q.includes('where did my money go')) {
      if (leakageItems.length === 0 && subscriptions.length === 0) {
        reply = "Great news! I scanned your recent ledger entries and couldn't find any critical leakage patterns. Keep adding daily logs to stay sharp.";
      } else {
        reply = `I identified bachelor savings opportunities:\n`;
        leakageItems.slice(0, 2).forEach(item => {
          reply += `• Small Leak: Spent ${fmt(item.monthlyCost)} monthly on ${item.merchant} (${item.category}). Extrapolates to ${fmt(item.yearlyCost)} annually. If redirected to mutual funds, this could be worth ${fmt(item.projectedWealth)} in 5 years.\n`;
        });
        subscriptions.filter(s => s.isLowUsage).forEach(sub => {
          reply += `• Subscription leak: ${sub.name} costing ${fmt(sub.amount)} with low usage (${sub.usage} times). Consider canceling.\n`;
        });
      }
    } else if (q.includes('health') || q.includes('score')) {
      reply = `Your Bachelor Health Score is ${healthScore}/100. Let's break it down:\n• Personal burn rate: ${fmt(totalMonthlySpend)} spent of ${fmt(salary)} salary.\n• Recommendation: Reduce food delivery and retail shopping to push your score above 80.`;
    } else if (q.includes('risk') || q.includes('safety')) {
      reply = `Your Risk Level: ${riskStatus.label}.\n${riskStatus.desc}\nYour personal commitments are ${fmt(totals.month * 0.25)} and liquidity is ${fmt(remainingSalary)}. Increasing emergency cash is our main safety toggle.`;
    } else if (q.includes('survival') || q.includes('how long') || q.includes('last this month')) {
      if (survivalData.isOverbudget) {
        reply = `⚠️ Survival Warning: Based on your current daily pace of ${fmt(survivalData.burnRatePerDay)}, you are projected to run out of salary on Day ${survivalData.runoutDay} (${survivalData.totalDays - survivalData.runoutDay} days early!). Your projected monthly spending is ${fmt(survivalData.projectedTotalSpend)} against ${fmt(salary)} salary. Freeze discretionary spending immediately.`;
      } else {
        reply = `You are on pace to survive comfortably this month! Your daily burn rate is ${fmt(survivalData.burnRatePerDay)}, leaving you with a projected surplus of ${fmt(salary - survivalData.projectedTotalSpend)} at month-end. Keep it up!`;
      }
    } else if (q.includes('split') || q.includes('roommate') || q.includes('owe')) {
      if (roommateDues.duesList.length === 0) {
        reply = "All roommates are perfectly squared up! No outstanding dues.";
      } else {
        reply = `Outstanding Roommate Balances:\n`;
        roommateDues.duesList.forEach(due => {
          reply += `• ${due.from} owes ${due.to} ${fmt(due.amount)}\n`;
        });
        if (currentUser) {
          reply += `Your net balance is ${fmt(roommateDues.balances[currentUser.name] || 0)}. If positive, you are owed money. If negative, you owe roommates.`;
        }
      }
    } else {
      reply = `I've analyzed your financial ledger. Based on your current monthly salary (${fmt(salary)}), you have logged ${fmt(totals.month)} this month, leaving ${fmt(remainingSalary)} remaining. Your current risk level is ${riskStatus.label} and score is ${healthScore}/100. Ask me about "spending leaks", "roommate split check", or "health score" for deep stats!`;
    }

    setChatMessages(prev => [
      ...prev,
      { id: Date.now(), sender: 'user', text: queryText },
      { id: Date.now() + 1, sender: 'coach', text: reply }
    ]);
    setChatInput('');
  };

  // --- Ingestion Simulated Parsers ---
  const handleCsvUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    setCsvFilePreview({
      name: file.name,
      size: (file.size / 1024).toFixed(1) + ' KB',
      rows: [
        { date: isoDate(new Date()), merchant: 'Zomato Delivery', amount: 480, category: 'Food', split: 'Yes' },
        { date: isoDate(new Date()), merchant: 'Uber India', amount: 150, category: 'Transport', split: 'No' },
        { date: isoDate(new Date()), merchant: 'Excitel WiFi', amount: 900, category: 'Bills', split: 'Yes' }
      ]
    });
  };

  const importCsvPreviewRows = async () => {
    if (!csvFilePreview || !session) return;
    const N = roommates.length || 1;
    const newRows = csvFilePreview.rows.map((row, idx) => ({
      user_id: session.user.id,
      room_id: row.split === 'Yes' ? currentRoomId : null,
      category: row.category,
      amount: row.amount,
      merchant: row.merchant,
      note: 'Imported from bank CSV statement [source:upload]',
      is_shared: row.split === 'Yes',
      logged_by: currentUser?.name || 'Roommate',
      date: row.date
    }));

    try {
      const { error } = await supabase.from('transactions').insert(newRows);
      if (error) {
        console.error('importCsvPreviewRows Supabase insert error details:', error);
        throw error;
      }
      showToast('success', `Imported ${newRows.length} transactions from statement!`);
      setCsvFilePreview(null);
      await fetchTransactions(session.user.id, currentRoomId);
    } catch (e) {
      showToast('error', 'Error importing transactions.');
    }
  };

  const handleReceiptOcrScan = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setOcrFile(file);
    setOcrScanning(true);
    setOcrResult(null);
    setTimeout(() => {
      setOcrScanning(false);
      setOcrResult({
        date: isoDate(new Date()),
        merchant: 'Decathlon Sports',
        amount: 2450,
        category: 'Shopping',
        confidence: '98%'
      });
    }, 2000);
  };

  const acceptOcrTransaction = async () => {
    if (!ocrResult || !session) return;
    const newTx = {
      user_id: session.user.id,
      room_id: null,
      category: ocrResult.category,
      amount: ocrResult.amount,
      merchant: ocrResult.merchant,
      note: 'Scanned receipt OCR validation [source:receipt]',
      is_shared: false,
      logged_by: currentUser?.name || 'Roommate',
      date: ocrResult.date
    };

    try {
      const { error } = await supabase.from('transactions').insert(newTx);
      if (error) {
        console.error('acceptOcrTransaction Supabase insert error details:', error);
        throw error;
      }
      showToast('success', 'Logged transaction from scanned receipt!');
      setOcrFile(null);
      setOcrResult(null);
      await fetchTransactions(session.user.id, currentRoomId);
    } catch (e) {
      showToast('error', 'Error logging receipt.');
    }
  };

  const triggerVoiceRecording = () => {
    if (voiceRecording) {
      setVoiceRecording(false);
      const command = "spent four hundred and fifty rupees on Zomato food delivery split with flatmates";
      setVoiceText(command);
      showToast('success', 'Voice captured.');
      setVoiceResult({
        date: isoDate(new Date()),
        merchant: 'Zomato Food Delivery',
        amount: 450,
        category: 'Food',
        isShared: true,
        summary: "Spent ₹450 on Food at Zomato (Split equally)"
      });
    } else {
      setVoiceText('');
      setVoiceResult(null);
      setVoiceRecording(true);
      setTimeout(() => {
        setVoiceRecording(false);
        triggerVoiceRecording();
      }, 3000);
    }
  };

  const acceptVoiceTransaction = async () => {
    if (!voiceResult || !session) return;
    const newTx = {
      user_id: session.user.id,
      room_id: voiceResult.isShared ? currentRoomId : null,
      category: voiceResult.category,
      amount: voiceResult.amount,
      merchant: voiceResult.merchant,
      note: 'Voice command transcription entry [source:voice]',
      is_shared: voiceResult.isShared,
      logged_by: currentUser?.name || 'Roommate',
      date: voiceResult.date
    };

    try {
      const { error } = await supabase.from('transactions').insert(newTx);
      if (error) {
        console.error('acceptVoiceTransaction Supabase insert error details:', error);
        throw error;
      }
      showToast('success', 'Logged transaction via voice entry command!');
      setVoiceText('');
      setVoiceResult(null);
      await fetchTransactions(session.user.id, currentRoomId);
    } catch (e) {
      showToast('error', 'Error logging voice command.');
    }
  };

  // --- Dues Clear Action ---
  const handleClearDuesDirect = async (from, to, amount) => {
    if (!session || !currentRoomId) return;
    if (!window.confirm(`Mark ₹${amount} dues payment as settled from ${from} to ${to}?`)) return;

    // Log an adjustment transaction:
    // Payer (from) logs a shared transaction with amount * N so they get credited
    // Wait! A cleaner way is to insert a specific settlement record or write an offset transaction.
    const N = roommates.length + 1;
    
    // Resolve which user logged it
    const adjustTx = {
      user_id: session.user.id,
      room_id: currentRoomId,
      category: 'Other',
      amount: amount * N, // mathematically offsets balances
      merchant: `Settle: ${from} to ${to}`,
      note: `Direct roommates dues settlement logged by ${currentUser.name} [source:manual]`,
      is_shared: true,
      logged_by: from, // credited to the person who paid
      date: isoDate(new Date())
    };

    try {
      const { error } = await supabase.from('transactions').insert(adjustTx);
      if (error) {
        console.error('handleClearDuesDirect Supabase insert error details:', error);
        throw error;
      }
      showToast('success', `Settled: Recorded ₹${amount} payment from ${from} to ${to}`);
      await fetchTransactions(session.user.id, currentRoomId);
    } catch (e) {
      showToast('error', 'Error logging settlement.');
    }
  };

  // Period-based Excel export
  const exportToExcelCustom = (period = 'all') => {
    const today = new Date();
    const todayStr = isoDate(today);

    // Get week start (Sunday)
    const weekStart = new Date(today);
    weekStart.setDate(today.getDate() - today.getDay());
    const weekStartStr = isoDate(weekStart);

    // Month start
    const monthStartStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-01`;

    // Year start
    const yearStartStr = `${today.getFullYear()}-01-01`;

    // Filter by personal or shared depending on active mode
    const isPersonal = analysisType === 'personal';
    const baseTxs = transactions.filter(t => (isPersonal ? !t.is_shared : t.is_shared) && t.category !== 'System');

    let filtered = baseTxs;
    let label = isPersonal ? 'Personal_All_Time' : 'Roommates_All_Time';
    const modeLabel = isPersonal ? 'Personal' : 'Roommates';
    
    if (period === 'today') {
      filtered = baseTxs.filter(t => t.date === todayStr);
      label = `${modeLabel}_Today_${todayStr}`;
    } else if (period === 'week') {
      filtered = baseTxs.filter(t => t.date >= weekStartStr && t.date <= todayStr);
      label = `${modeLabel}_Week_${weekStartStr}_to_${todayStr}`;
    } else if (period === 'month') {
      filtered = baseTxs.filter(t => t.date >= monthStartStr && t.date <= todayStr);
      const monthName = today.toLocaleString('default', { month: 'long' });
      label = `${modeLabel}_Month_${monthName}_${today.getFullYear()}`;
    } else if (period === 'year') {
      filtered = baseTxs.filter(t => t.date >= yearStartStr && t.date <= todayStr);
      label = `${modeLabel}_Year_${today.getFullYear()}`;
    }

    if (filtered.length === 0) {
      showToast('error', `No transactions found for: ${period}`);
      return;
    }

    // Sort by date descending
    filtered.sort((a, b) => b.date.localeCompare(a.date));

    // Construct spreadsheet rows
    let roll = [];
    if (isPersonal) {
      roll = filtered.map(t => ({
        Date: t.date,
        Day: new Date(t.date).toLocaleDateString('en-IN', { weekday: 'long' }),
        Category: t.category,
        Merchant: t.merchant,
        Note: t.note || '',
        'Amount (₹)': t.amount,
        Source: t.source,
      }));
      
      const total = filtered.reduce((s, t) => s + Number(t.amount), 0);
      roll.push({});
      roll.push({ Date: 'TOTAL', 'Amount (₹)': total });
    } else {
      // Roommates Split Mode
      roll = filtered.map(t => ({
        Date: t.date,
        Day: new Date(t.date).toLocaleDateString('en-IN', { weekday: 'long' }),
        Category: t.category,
        Merchant: t.merchant,
        Note: t.note || '',
        'Paid By': t.logged_by,
        'Amount (₹)': t.amount,
        Source: t.source,
      }));

      const total = filtered.reduce((s, t) => s + Number(t.amount), 0);
      roll.push({});
      roll.push({ Date: 'GRAND TOTAL', 'Amount (₹)': total });

      // Calculate total paid by each person
      const totalsByPerson = {};
      
      // Initialize with all roommates + current user
      const allNames = [currentUser?.name || 'Me', ...roommates.map(r => r.name)];
      allNames.forEach(name => {
        totalsByPerson[name] = 0;
      });

      filtered.forEach(t => {
        const name = t.logged_by;
        totalsByPerson[name] = (totalsByPerson[name] || 0) + Number(t.amount);
      });

      roll.push({});
      roll.push({ Date: 'BREAKDOWN BY PERSON:' });
      
      Object.entries(totalsByPerson).forEach(([name, amt]) => {
        roll.push({ Date: name, 'Amount (₹)': amt });
      });

      const count = allNames.length || 1;
      const share = total / count;
      roll.push({});
      roll.push({ Date: `Individual Share (divided by ${count})`, 'Amount (₹)': share });
    }

    const ws = XLSX.utils.json_to_sheet(roll);
    
    // Set column widths
    if (isPersonal) {
      ws['!cols'] = [{ wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 20 }, { wch: 20 }, { wch: 12 }, { wch: 12 }];
    } else {
      ws['!cols'] = [{ wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 20 }, { wch: 20 }, { wch: 15 }, { wch: 12 }, { wch: 12 }];
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, label.substring(0, 31));
    XLSX.writeFile(wb, `FinSense_${label}.xlsx`);
    showToast('success', `Downloaded: ${filtered.length} entries for ${period}`);
  };

  // Period-based PDF export
  const exportToPDFCustom = (period = 'all') => {
    const today = new Date();
    const todayStr = isoDate(today);

    // Get week start (Sunday)
    const weekStart = new Date(today);
    weekStart.setDate(today.getDate() - today.getDay());
    const weekStartStr = isoDate(weekStart);

    // Month start
    const monthStartStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-01`;

    // Year start
    const yearStartStr = `${today.getFullYear()}-01-01`;

    const isPersonal = analysisType === 'personal';
    const baseTxs = transactions.filter(t => (isPersonal ? !t.is_shared : t.is_shared) && t.category !== 'System');

    let filtered = baseTxs;
    const modeLabel = isPersonal ? 'Personal' : 'Roommates';
    let periodLabel = 'All Time';
    
    if (period === 'today') {
      filtered = baseTxs.filter(t => t.date === todayStr);
      periodLabel = `Today (${todayStr})`;
    } else if (period === 'week') {
      filtered = baseTxs.filter(t => t.date >= weekStartStr && t.date <= todayStr);
      periodLabel = `This Week (${weekStartStr} to ${todayStr})`;
    } else if (period === 'month') {
      filtered = baseTxs.filter(t => t.date >= monthStartStr && t.date <= todayStr);
      const monthName = today.toLocaleString('default', { month: 'long' });
      periodLabel = `This Month (${monthName} ${today.getFullYear()})`;
    } else if (period === 'year') {
      filtered = baseTxs.filter(t => t.date >= yearStartStr && t.date <= todayStr);
      periodLabel = `This Year (${today.getFullYear()})`;
    }

    if (!isPersonal && shouldInjectRent && period !== 'today') {
      filtered = [
        ...filtered,
        {
          id: 'virtual-rent',
          user_id: roomAdminId || '',
          room_id: currentRoomId,
          category: 'Rent',
          amount: activeRentAmount,
          merchant: 'Flat Owner (Auto Rent Split)',
          note: `Auto split injected: ${fmt(activeRentAmount)}`,
          is_shared: true,
          logged_by: 'Auto Split',
          date: todayStr
        }
      ];
    }

    if (filtered.length === 0) {
      showToast('error', `No transactions found for: ${period}`);
      return;
    }

    // Sort by date descending
    filtered.sort((a, b) => b.date.localeCompare(a.date));

    const total = filtered.reduce((s, t) => s + Number(t.amount), 0);

    // Open print window
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      showToast('error', 'Popup blocker prevented report window. Please allow popups.');
      return;
    }

    // Start HTML
    let html = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>SmartSave Expense Report - ${modeLabel} (${periodLabel})</title>
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
          
          body {
            font-family: 'Outfit', sans-serif;
            background-color: #FBF8F0;
            color: #1E2229;
            margin: 0;
            padding: 40px;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
          }
          
          .header {
            border-bottom: 2px solid #1E2229;
            padding-bottom: 20px;
            margin-bottom: 30px;
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
          }
          
          .title {
            margin: 0;
            font-size: 26px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: -0.5px;
          }
          
          .meta {
            font-size: 11px;
            text-transform: uppercase;
            font-weight: 600;
            color: #6B7280;
            letter-spacing: 0.5px;
          }

          .date-stamp {
            font-family: 'JetBrains Mono', monospace;
            font-size: 11px;
            text-align: right;
          }
          
          .summary-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            margin-bottom: 30px;
          }
          
          .summary-card {
            background-color: #FAF6ED;
            border: 1px solid #1E2229;
            padding: 15px;
            border-radius: 4px;
            box-shadow: 2px 2px 0px #1E2229;
          }
          
          .summary-card-title {
            font-size: 10px;
            font-weight: 800;
            text-transform: uppercase;
            color: #6B7280;
            letter-spacing: 0.5px;
          }
          
          .summary-card-value {
            font-family: 'JetBrains Mono', monospace;
            font-size: 18px;
            font-weight: 700;
            margin-top: 5px;
          }
          
          .section-title {
            font-size: 14px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 12px;
            border-left: 4px solid #1E2229;
            padding-left: 8px;
          }
          
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 40px;
          }
          
          th {
            background-color: #1E2229;
            color: #FBF8F0;
            text-align: left;
            padding: 8px 12px;
            font-size: 10px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            font-weight: 800;
          }
          
          td {
            padding: 10px 12px;
            font-size: 12px;
            border-bottom: 1px solid #E5E7EB;
          }
          
          tr:last-child td {
            border-bottom: 2px solid #1E2229;
          }
          
          .text-right {
            text-align: right;
          }
          
          .mono {
            font-family: 'JetBrains Mono', monospace;
          }

          .badge {
            display: inline-block;
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 9px;
            font-weight: 800;
            text-transform: uppercase;
            border: 1px solid currentColor;
          }

          .badge-recv {
            color: #065F46;
            background-color: #D1FAE5;
          }

          .badge-pay {
            color: #991B1B;
            background-color: #FEE2E2;
          }

          .badge-settled {
            color: #374151;
            background-color: #F3F4F6;
          }
          
          .footer {
            margin-top: 50px;
            border-top: 1px solid #E5E7EB;
            padding-top: 15px;
            text-align: center;
            font-size: 10px;
            color: #9CA3AF;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            font-weight: 600;
          }
          
          @media print {
            body {
              padding: 0;
              background-color: #FFFFFF;
            }
            .summary-card {
              box-shadow: none;
              background-color: #FFFFFF;
            }
          }
        </style>
      </head>
      <body>
        <div class="header">
          <div>
            <span class="meta">SmartSave Financial Statement</span>
            <h1 class="title">${modeLabel} Report</h1>
            <span class="meta" style="color: #1E2229;">Period: ${periodLabel}</span>
          </div>
          <div class="date-stamp">
            Generated: ${new Date().toLocaleString('en-IN')}<br>
            Mode: ${isPersonal ? 'Bachelor Personal' : 'Roommates Split'}
          </div>
        </div>
    `;

    if (isPersonal) {
      const remaining = salary - total;
      const savingsPct = salary > 0 ? Math.round((remaining / salary) * 100) : 0;

      // Personal Summary Cards
      html += `
        <div class="summary-grid">
          <div class="summary-card">
            <div class="summary-card-title">Monthly Income</div>
            <div class="summary-card-value">₹${salary.toLocaleString('en-IN')}</div>
          </div>
          <div class="summary-card">
            <div class="summary-card-title">Total Expense</div>
            <div class="summary-card-value">₹${total.toLocaleString('en-IN')}</div>
          </div>
          <div class="summary-card">
            <div class="summary-card-title">Remaining Balance</div>
            <div class="summary-card-value">₹${remaining.toLocaleString('en-IN')}</div>
          </div>
          <div class="summary-card">
            <div class="summary-card-title">Savings Rate</div>
            <div class="summary-card-value">${savingsPct}%</div>
          </div>
        </div>

        <div class="section-title">Expense Itemization</div>
        <table>
          <thead>
            <tr>
              <th style="width: 15%;">Date</th>
              <th style="width: 15%;">Day</th>
              <th style="width: 20%;">Category</th>
              <th style="width: 35%;">Description (Merchant/Note)</th>
              <th style="width: 15%; text-align: right;">Amount</th>
            </tr>
          </thead>
          <tbody>
      `;

      filtered.forEach(t => {
        const dayName = new Date(t.date).toLocaleDateString('en-IN', { weekday: 'long' });
        const desc = t.note ? `${t.merchant} (${t.note})` : t.merchant;
        html += `
          <tr>
            <td class="mono">${t.date}</td>
            <td>${dayName}</td>
            <td>${t.category}</td>
            <td>${desc}</td>
            <td class="text-right mono font-bold">₹${Number(t.amount).toLocaleString('en-IN')}</td>
          </tr>
        `;
      });

      html += `
          </tbody>
        </table>
      `;

    } else {
      // Roommates Split Mode
      const count = roommates.length + 1;
      const share = total / count;

      const balances = {};
      const totalPaidMap = {};
      
      const allNames = [currentUser?.name || 'Me', ...roommates.map(r => r.name)].filter(Boolean);
      allNames.forEach(name => {
        balances[name] = 0;
        totalPaidMap[name] = 0;
      });

      filtered.forEach(tx => {
        if (tx.id === 'virtual-rent') return;
        const payer = tx.logged_by;
        const txShare = Number(tx.amount) / count;
        
        if (totalPaidMap[payer] !== undefined) {
          totalPaidMap[payer] += Number(tx.amount);
        }
        
        allNames.forEach(name => {
          if (balances[name] !== undefined) {
            if (name === payer) {
              balances[name] += Number(tx.amount) - txShare;
            } else {
              balances[name] -= txShare;
            }
          }
        });
      });

      // Calculate dues List for this specific filtered period
      const debtors = [];
      const creditors = [];
      Object.entries(balances).forEach(([name, bal]) => {
        if (bal < -0.99) debtors.push({ name, bal: -bal });
        if (bal > 0.99) creditors.push({ name, bal });
      });

      const periodDuesList = [];
      let dIdx = 0;
      let cIdx = 0;
      while (dIdx < debtors.length && cIdx < creditors.length) {
        const debtor = debtors[dIdx];
        const creditor = creditors[cIdx];
        const transfer = Math.min(debtor.bal, creditor.bal);
        periodDuesList.push({
          from: debtor.name,
          to: creditor.name,
          amount: Math.round(transfer)
        });
        debtor.bal -= transfer;
        creditor.bal -= transfer;
        if (debtor.bal <= 1) dIdx++;
        if (creditor.bal <= 1) cIdx++;
      }

      // Roommate Summary Cards
      html += `
        <div class="summary-grid">
          <div class="summary-card">
            <div class="summary-card-title">Total Room Expense</div>
            <div class="summary-card-value">₹${Math.round(total).toLocaleString('en-IN')}</div>
          </div>
          <div class="summary-card">
            <div class="summary-card-title">Total Roommates</div>
            <div class="summary-card-value">${count} members</div>
          </div>
          <div class="summary-card">
            <div class="summary-card-title">Per Person Share</div>
            <div class="summary-card-value">₹${Math.round(share).toLocaleString('en-IN')}</div>
          </div>
          <div class="summary-card">
            <div class="summary-card-title">Pending Settlements</div>
            <div class="summary-card-value">${periodDuesList.length} transfers</div>
          </div>
        </div>

        <div class="section-title">Roommate Settlement Status</div>
        <table style="margin-bottom: 25px;">
          <thead>
            <tr>
              <th>Member</th>
              <th style="text-align: right;">Total Paid</th>
              <th style="text-align: right;">Share Amount</th>
              <th style="text-align: right;">Pay</th>
              <th style="text-align: right;">Receive</th>
              <th style="text-align: right;">Status</th>
            </tr>
          </thead>
          <tbody>
      `;

      const otherTotal = filtered.filter(tx => tx.id !== 'virtual-rent').reduce((s, tx) => s + Number(tx.amount), 0);
      const otherShare = otherTotal / count;
      const rentShare = shouldInjectRent && period !== 'today' ? (activeRentAmount / count) : 0;
      const totalShare = otherShare + rentShare;

      allNames.forEach(name => {
        const paid = totalPaidMap[name] || 0;
        const bal = paid - otherShare; // Settlement balance (excluding rent)
        const pays = bal < 0 ? Math.round(-bal) : 0;
        const recvs = bal > 0 ? Math.round(bal) : 0;
        
        let badgeClass = 'badge-settled';
        let statusText = 'Settled';
        if (bal > 0.99) {
          badgeClass = 'badge-recv';
          statusText = 'Receive';
        } else if (bal < -0.99) {
          badgeClass = 'badge-pay';
          statusText = 'Pay';
        }

        html += `
          <tr>
            <td><strong>${name}</strong> ${name === currentUser?.name ? '(You)' : ''}</td>
            <td class="text-right mono">₹${Math.round(paid).toLocaleString('en-IN')}</td>
            <td class="text-right mono">₹${Math.round(totalShare).toLocaleString('en-IN')}</td>
            <td class="text-right mono text-red-700">${pays > 0 ? `₹${pays.toLocaleString('en-IN')}` : '-'}</td>
            <td class="text-right mono text-emerald-800">${recvs > 0 ? `₹${recvs.toLocaleString('en-IN')}` : '-'}</td>
            <td class="text-right"><span class="badge ${badgeClass}">${statusText}</span></td>
          </tr>
        `;
      });

      html += `
          </tbody>
        </table>

        <div class="section-title">Calculated Settlements Needed</div>
        <table style="margin-bottom: 30px;">
          <thead>
            <tr>
              <th>From (Debtor)</th>
              <th>To (Creditor)</th>
              <th style="text-align: right;">Settlement Amount</th>
            </tr>
          </thead>
          <tbody>
      `;

      if (periodDuesList.length === 0) {
        html += `
          <tr>
            <td colspan="3" style="text-align: center; color: #6B7280; font-style: italic;">Everyone is squared up! No settlements needed for this period.</td>
          </tr>
        `;
      } else {
        periodDuesList.forEach(due => {
          html += `
            <tr>
              <td><span style="color: #B91C1C; font-weight: 600;">${due.from}</span></td>
              <td><span style="color: #047857; font-weight: 600;">${due.to}</span></td>
              <td class="text-right mono font-bold">₹${due.amount.toLocaleString('en-IN')}</td>
            </tr>
          `;
        });
      }

      html += `
          </tbody>
        </table>

        <div class="section-title">Flat Shared Expenses Ledger</div>
        <table>
          <thead>
            <tr>
              <th style="width: 12%;">Date</th>
              <th style="width: 12%;">Day</th>
              <th style="width: 18%;">Buyer Name</th>
              <th style="width: 18%;">Category</th>
              <th style="width: 25%;">Item Purchased (Merchant/Notes)</th>
              <th style="width: 15%; text-align: right;">Amount</th>
            </tr>
          </thead>
          <tbody>
      `;

      filtered.forEach(t => {
        const dayName = new Date(t.date).toLocaleDateString('en-IN', { weekday: 'long' });
        const desc = t.note ? `${t.merchant} (${t.note})` : t.merchant;
        html += `
          <tr>
            <td class="mono">${t.date}</td>
            <td>${dayName}</td>
            <td><strong>${t.logged_by}</strong></td>
            <td>${t.category}</td>
            <td>${desc}</td>
            <td class="text-right mono font-bold">₹${Number(t.amount).toLocaleString('en-IN')}</td>
          </tr>
        `;
      });

      html += `
          </tbody>
        </table>
      `;
    }

    // End HTML
    html += `
        <div class="footer">
          SmartSave App &bull; Bachelor & Roommate Expense Management System &bull; &copy; ${today.getFullYear()}
        </div>
        <script>
          window.onload = function() {
            setTimeout(function() {
              window.print();
            }, 500);
          }
        </script>
      </body>
      </html>
    `;

    printWindow.document.write(html);
    printWindow.document.close();
    showToast('success', `PDF Report generated for ${period}`);
  };

  const [exportMenuOpen, setExportMenuOpen] = React.useState(false);
  const [pdfMenuOpen, setPdfMenuOpen] = useState(false);
  const [ledgerPeriod, setLedgerPeriod] = React.useState('all'); // 'today'|'week'|'month'|'year'|'all'

  // Dynamic styles and styling bindings
  const fontStyle = useMemo(() => {
    let family = 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    if (fontFamily === 'serif') {
      family = 'Georgia, Cambria, "Times New Roman", Times, serif';
    } else if (fontFamily === 'mono') {
      family = 'ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace';
    }

    let size = '14px';
    let labelSize = '11px';
    let titleSize = '16px';
    if (fontSize === 'small') {
      size = '12px';
      labelSize = '9px';
      titleSize = '14px';
    } else if (fontSize === 'large') {
      size = '16px';
      labelSize = '12px';
      titleSize = '18px';
    }

    let tracking = 'normal';
    if (layoutDensity === 'tight') {
      tracking = '-0.015em';
    }

    return {
      family,
      size,
      labelSize,
      titleSize,
      tracking
    };
  }, [fontFamily, fontSize, layoutDensity]);

  const headerContent = useMemo(() => {
    switch (activeTab) {
      case 'dashboard':
        return {
          label: 'Dashboard Overview',
          title: analysisType === 'personal' ? `${currentUser?.name || 'My'}'s Personal Dashboard` : 'Roommate Splits Dashboard'
        };
      case 'transactions':
        return {
          label: 'Daily Expenses',
          title: analysisType === 'personal' ? `${currentUser?.name || 'My'}'s Personal Expenses` : 'Roommates Daily Splits'
        };
      case 'waste':
        return {
          label: 'Personal Leakage',
          title: 'Daily Leakage Finder'
        };
      case 'goals':
        return {
          label: 'Goals & Dues',
          title: 'Goals & Roommate Dues'
        };
      case 'coach':
        return {
          label: 'AI Coach',
          title: 'AI Financial Coach & Forecasts'
        };
      case 'ingestion':
        return {
          label: 'Statement / Voice Hub',
          title: 'Transaction Ingestion Hub'
        };
      default:
        return {
          label: 'Daily Expenses',
          title: 'Daily Expenses Ledger'
        };
    }
  }, [activeTab, analysisType, currentUser]);


  const currentTheme = THEMES[themeMode] || THEMES.beige;

  const rootStyles = {
    '--paper': currentTheme.paper,
    '--paper-deep': currentTheme.paperDeep,
    '--card': currentTheme.card,
    '--ink': currentTheme.ink,
    '--ink-soft': currentTheme.inkSoft,
    '--rule': currentTheme.rule,
    '--stamp': currentTheme.stamp,
    '--positive': currentTheme.positive,
    '--border': currentTheme.border,
    '--accent': currentTheme.accent,
    background: currentTheme.paper,
    borderColor: currentTheme.rule
  };

  // --- Auth Screen Overlay ---
  if (!session) {
    return (
      <div className="ledger-app min-h-screen flex items-center justify-center p-4 transition-all duration-300" style={rootStyles}>
        <style>{`
          .ledger-app {
            font-family: ${fontStyle.family};
            font-size: ${fontStyle.size};
            letter-spacing: ${fontStyle.tracking};
            line-height: 1.35;
            color: var(--ink);
          }
          .auth-card {
            background: var(--card);
            border: 1px solid var(--rule);
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
          }
          .ledger-input-box {
            border: 1px solid var(--rule);
            background: var(--card);
            border-radius: 4px;
            padding: 8px 12px;
            font-size: 13px;
            color: var(--ink);
            outline: none;
            width: 100%;
          }
          .ledger-input-box:focus {
            border-color: var(--ink);
            box-shadow: 0 0 0 1px var(--ink);
          }
          .btn-vintage-ink {
            background: var(--ink);
            color: var(--card);
            font-weight: 600;
            font-size: 13px;
            padding: 9px 16px;
            border-radius: 4px;
            border: none;
            cursor: pointer;
            width: 100%;
            transition: opacity 0.15s;
          }
          .btn-vintage-ink:hover {
            opacity: 0.95;
          }
        `}</style>
        
        <div className="auth-card max-w-sm w-full rounded-lg p-6 space-y-6">
          <div className="text-center space-y-1">
            <h1 className="text-xl font-bold text-[var(--ink)]">FinSense AI</h1>
            <p className="text-[10px] uppercase tracking-wider text-[var(--ink-soft)] font-bold">Collaborative Roommate Ledger</p>
          </div>

          <form onSubmit={handleAuthSubmit} className="space-y-4">
            {isRegistering && (
              <div className="flex flex-col">
                <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Your Full Name</label>
                <input
                  type="text"
                  placeholder="e.g. Siva"
                  className="ledger-input-box"
                  value={authName}
                  onChange={e => setAuthName(e.target.value)}
                  required
                />
              </div>
            )}

            <div className="flex flex-col">
              <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Email Address</label>
              <input
                type="email"
                placeholder="siva@bachelor.com"
                className="ledger-input-box"
                value={authEmail}
                onChange={e => setAuthEmail(e.target.value)}
                required
              />
            </div>

            <div className="flex flex-col">
              <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Password</label>
              <input
                type="password"
                placeholder="••••••••"
                className="ledger-input-box"
                value={authPassword}
                onChange={e => setAuthPassword(e.target.value)}
                required
              />
            </div>

            {errorMsg && (
              <p className="text-xs text-[var(--stamp)] font-bold text-center">{errorMsg}</p>
            )}

            <button type="submit" className="btn-vintage-ink flex items-center justify-center gap-1.5" disabled={authLoading}>
              {authLoading ? "Processing..." : isRegistering ? "Create Account" : "Sign In"}
            </button>
          </form>

          <div className="text-center">
            <button
              onClick={() => {
                setIsRegistering(!isRegistering);
                setErrorMsg(null);
              }}
              className="text-xs underline font-bold"
              style={{ color: 'var(--ink)' }}
            >
              {isRegistering ? "Already have an account? Sign In" : "Need an account? Sign Up"}
            </button>
          </div>
        </div>
      </div>
    );
  }



  return (
    <div className="ledger-app min-h-screen md:min-h-[660px] flex flex-col rounded-lg overflow-hidden shadow-xl border text-slate-800 transition-all duration-300 w-full md:w-auto md:max-w-7xl md:mx-auto md:my-8" style={rootStyles}>
      
      {/* Dynamic Styling injections */}
      <style>{`
        .ledger-app {
          font-family: ${fontStyle.family};
          font-size: ${fontStyle.size};
          letter-spacing: ${fontStyle.tracking};
          line-height: 1.35;
          color: var(--ink);
        }
        .ledger-display {
          font-family: ${fontStyle.family};
          letter-spacing: -0.02em;
          font-weight: 700;
        }
        .ledger-mono {
          font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
          font-variant-numeric: tabular-nums;
          letter-spacing: -0.015em;
        }
        .blur-card {
          background: var(--card);
          border: 1px solid var(--rule);
          box-shadow: 0 1px 3px rgba(30,42,68,0.05);
        }
        .perforation {
          width: 16px;
          background-image: radial-gradient(circle at 8px center, var(--paper) 4px, transparent 4.5px);
          background-size: 16px 22px;
          background-repeat: repeat-y;
          background-color: var(--paper-deep);
          flex-shrink: 0;
        }
        .tab-btn-custom {
          font-family: ${fontStyle.family};
          font-size: 12px;
          letter-spacing: normal;
          transition: all 0.15s ease-in-out;
          border-right: 1px solid var(--rule);
          border-bottom: 1px solid var(--rule);
          color: var(--ink-soft);
        }
        .tab-btn-custom:hover {
          background: rgba(30,42,68,0.04);
        }
        .tab-btn-custom.active {
          background: var(--card);
          color: var(--ink);
          font-weight: 700;
          border-right: none;
          border-bottom: none;
        }
        .stamp-badge {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          border: 2px solid var(--stamp);
          color: var(--stamp);
          font-weight: 700;
          font-size: 11px;
          letter-spacing: -0.01em;
          text-transform: uppercase;
          padding: 3px 8px;
          border-radius: 3px;
          transform: rotate(-2deg);
        }
        .ok-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          color: var(--positive);
          font-weight: 700;
          font-size: 11px;
          letter-spacing: -0.01em;
          text-transform: uppercase;
          padding: 3px 8px;
          border: 2px solid var(--positive);
          border-radius: 3px;
          transform: rotate(1deg);
        }
        .ledger-input-box {
          border: 1px solid var(--rule);
          background: var(--card);
          border-radius: 3px;
          padding: 6px 9px;
          font-size: 13px;
          color: var(--ink);
          outline: none;
        }
        .ledger-input-box:focus {
          border-color: var(--ink);
          box-shadow: 0 0 0 1px var(--ink);
        }
        .btn-vintage-ink {
          background: var(--ink);
          color: var(--card);
          font-weight: 600;
          font-size: 12px;
          padding: 7px 14px;
          border-radius: 3px;
          border: none;
          cursor: pointer;
          transition: opacity 0.15s;
        }
        .btn-vintage-ink:hover {
          opacity: 0.9;
        }
        .btn-vintage-outline {
          background: transparent;
          border: 1px solid var(--ink);
          color: var(--ink);
          font-weight: 600;
          font-size: 12px;
          padding: 6px 12px;
          border-radius: 3px;
          cursor: pointer;
        }
        .scanner-beam {
          animation: scan 2s linear infinite;
        }
        @keyframes scan {
          0% { top: 0%; opacity: 0.8; }
          50% { top: 100%; opacity: 0.8; }
          100% { top: 0%; opacity: 0.8; }
        }
      `}</style>

      {/* Mobile Top App Bar */}
      <header className="md:hidden flex items-center justify-between p-4 border-b bg-[var(--paper-deep)] sticky top-0 z-30" style={{ borderColor: 'var(--rule)' }}>
        <div className="flex items-center gap-2">
          <button 
            onClick={() => setIsMobileMenuOpen(true)}
            className="p-1.5 rounded hover:bg-slate-800/10 text-[var(--ink)] focus:outline-none"
            aria-label="Open menu"
          >
            <Icons.Menu className="w-5 h-5" />
          </button>
          <div className="w-6 h-6 rounded flex items-center justify-center font-bold text-xs text-[var(--card)]" style={{ background: 'var(--ink)' }}>
            F
          </div>
          <span className="font-bold text-xs tracking-tight text-slate-900">FinSense AI</span>
        </div>
        <div className="flex items-center gap-2">
          {currentRoomCode && (
            <span className="text-[8px] uppercase font-bold tracking-wider text-emerald-800 bg-emerald-600/10 px-1.5 py-0.5 rounded border border-emerald-500/20">
              Flat: {currentRoomCode}
            </span>
          )}
          <div className="w-6 h-6 rounded-full bg-[var(--ink)] text-[var(--card)] flex items-center justify-center text-[10px] font-bold uppercase">
            {currentUser?.name?.slice(0, 2) || 'U'}
          </div>
        </div>
      </header>

      {/* Main Inner Flex Wrapper */}
      <div className="flex flex-col md:flex-row flex-1 min-h-0">
        
        {/* Mobile menu backdrop overlay */}
        {isMobileMenuOpen && (
          <div 
            className="fixed inset-0 bg-black/40 backdrop-blur-xs z-40 md:hidden transition-opacity duration-300"
            onClick={() => setIsMobileMenuOpen(false)}
          />
        )}
      
        {/* Sidebar Navigation */}
        <aside className={`fixed inset-y-0 left-0 z-50 w-64 md:static md:w-64 flex-shrink-0 flex flex-col justify-between border-r transform transition-transform duration-300 ease-in-out md:translate-x-0 ${
          isMobileMenuOpen ? 'translate-x-0' : '-translate-x-full'
        }`} style={{ borderColor: 'var(--rule)', background: 'var(--paper-deep)' }}>
        <div>
          {/* Logo */}
          <div className="p-5 border-b flex items-center gap-2" style={{ borderColor: 'var(--rule)' }}>
            <div className="w-7 h-7 rounded flex items-center justify-center font-bold text-base text-[var(--card)]" style={{ background: 'var(--ink)' }}>
              F
            </div>
            <div>
              <h1 className="font-bold text-sm tracking-tight text-slate-900">FinSense AI</h1>
              <p className="text-[9px] uppercase tracking-wider font-semibold" style={{ color: 'var(--ink-soft)' }}>Live Flat roommate splits</p>
            </div>
          </div>

          {/* Active Profile Info */}
          <div className="p-4 border-b space-y-2.5" style={{ borderColor: 'var(--rule)' }}>
            <div className="flex items-center justify-between text-[10px] uppercase font-bold tracking-wider" style={{ color: 'var(--ink-soft)' }}>
              <span><Icons.User className="w-3.5 h-3.5 mr-1 inline text-[var(--ink-soft)]" /> profile info</span>
            </div>
            <div className="bg-white/40 p-2.5 rounded border border-dashed text-xs space-y-1.5" style={{ borderColor: 'var(--rule)' }}>
              <p className="font-bold text-[var(--ink)] truncate">User: {currentUser?.name}</p>
              <p className="text-[10px] text-slate-500 truncate">{currentUser?.email}</p>
              {/* Only show Flat Code + Add Flatmate after a flat is created/joined */}
              {roomMembershipStatus !== 'none' && (
                <div className="pt-1.5 border-t border-dashed flex justify-between items-center text-[10px]" style={{ borderColor: 'var(--rule)' }}>
                  <span className="font-bold uppercase text-[9px] text-emerald-800 bg-emerald-600/10 px-1 py-0.2 rounded border border-emerald-500/20">
                    Flat Code: {currentRoomCode}
                  </span>
                  {roomMembershipStatus === 'accepted' && (
                    <button 
                      onClick={() => {
                        setAuthForm({ name: '', email: '' });
                        setIsAuthOpen(true);
                      }}
                      className="underline hover:text-slate-900 font-bold flex items-center gap-0.5"
                      style={{ color: 'var(--ink)' }}
                    >
                      <Icons.AddUser /> Add Flatmate
                    </button>
                  )}
                </div>
              )}
            </div>
          </div>

          {/* Navigation vertical list with SVG Icons */}
          <nav className="flex flex-col text-xs font-semibold">
            <button
              onClick={() => { setActiveTab('dashboard'); setIsMobileMenuOpen(false); }}
              className={`tab-btn-custom w-full text-left py-3 px-4 flex items-center justify-between border-b ${activeTab === 'dashboard' ? 'active' : ''}`}
              style={{ borderColor: 'var(--rule)' }}
            >
              <span className="flex items-center"><Icons.Dashboard /> Dashboard Overview</span>
            </button>
            <button
              onClick={() => { setActiveTab('transactions'); setIsMobileMenuOpen(false); }}
              className={`tab-btn-custom w-full text-left py-3 px-4 flex items-center justify-between border-b ${activeTab === 'transactions' ? 'active' : ''}`}
              style={{ borderColor: 'var(--rule)' }}
            >
              <span className="flex items-center"><Icons.Ledger /> Daily Expenses</span>
            </button>
            <button
              onClick={() => { setActiveTab('waste'); setAnalysisType('personal'); setIsMobileMenuOpen(false); }}
              className={`tab-btn-custom w-full text-left py-3 px-4 flex items-center justify-between border-b ${activeTab === 'waste' ? 'active' : ''}`}
              style={{ borderColor: 'var(--rule)' }}
            >
              <span className="flex items-center"><Icons.Leakage /> Personal Leakage</span>
              {leakageItems.length + subscriptions.filter(s => s.isLowUsage).length > 0 && (
                <span className="w-1.5 h-1.5 rounded-full bg-[var(--stamp)]" />
              )}
            </button>
            <button
              onClick={() => { setActiveTab('goals'); setAnalysisType('personal'); setIsMobileMenuOpen(false); }}
              className={`tab-btn-custom w-full text-left py-3 px-4 flex items-center justify-between border-b ${activeTab === 'goals' ? 'active' : ''}`}
              style={{ borderColor: 'var(--rule)' }}
            >
              <span className="flex items-center"><Icons.Goals /> Goals & room dues</span>
              {roommateDues.duesList.length > 0 && (
                <span className="text-[9px] px-1 py-0.2 rounded bg-red-600/10 text-red-800 font-bold border border-red-500/20">{roommateDues.duesList.length}</span>
              )}
            </button>
            <button
              onClick={() => { setActiveTab('coach'); setAnalysisType('personal'); setIsMobileMenuOpen(false); }}
              className={`tab-btn-custom w-full text-left py-3 px-4 flex items-center justify-between border-b ${activeTab === 'coach' ? 'active' : ''}`}
              style={{ borderColor: 'var(--rule)' }}
            >
              <span className="flex items-center"><Icons.Coach /> AI Coach & Forecasting</span>
            </button>
            <button
              onClick={() => { setActiveTab('ingestion'); setAnalysisType('personal'); setIsMobileMenuOpen(false); }}
              className={`tab-btn-custom w-full text-left py-3 px-4 flex items-center justify-between border-b ${activeTab === 'ingestion' ? 'active' : ''}`}
              style={{ borderColor: 'var(--rule)' }}
            >
              <span className="flex items-center"><Icons.Ingestion /> Statement / Voice Hub</span>
            </button>
          </nav>
        </div>

        {/* Bottom Action Buttons */}
        <div className="p-4 border-t flex flex-col gap-1.5" style={{ borderColor: 'var(--rule)', background: 'var(--paper-deep)' }}>
          <button onClick={() => { seedDemoData(); setIsMobileMenuOpen(false); }} className="w-full py-1.5 bg-[#5A6483] hover:bg-[#4E5672] text-[#FBF8F0] text-[9px] uppercase font-bold tracking-wider rounded shadow-sm">
            <Icons.Seed className="w-3.5 h-3.5 mr-1 inline" /> Seed Flat Data
          </button>
          <button onClick={() => { clearAllData(); setIsMobileMenuOpen(false); }} className="w-full py-1 border border-[var(--stamp)] text-[var(--stamp)] hover:bg-red-50/20 text-[9px] uppercase font-bold tracking-wider rounded">
            <Icons.Broom className="w-3.5 h-3.5 mr-1 inline" /> Reset My Data
          </button>
          <button onClick={() => { supabase.auth.signOut(); setIsMobileMenuOpen(false); }} className="w-full py-1 border border-slate-500 text-slate-500 hover:bg-slate-50/20 text-[9px] uppercase font-bold tracking-wider rounded">
            Sign Out Account
          </button>
        </div>
      </aside>

      {/* Perforation Divider */}
      <div className="perforation hidden md:block" />

      {/* Main Content */}
      <main className="flex-1 p-5 md:p-8 flex flex-col justify-between bg-[var(--card)] min-w-0">
        
        {/* Header & Mode Switcher */}
          <div className="flex flex-col md:flex-row md:items-center justify-between pb-4 mb-6 border-b gap-4" style={{ borderColor: 'var(--rule)' }}>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs uppercase font-bold tracking-wider" style={{ color: 'var(--ink-soft)' }}>{headerContent.label}</span>
                {currentRoomCode && (
                  <span className="text-[9px] uppercase font-bold tracking-wider text-emerald-800 bg-emerald-600/10 px-1.5 py-0.2 rounded border border-emerald-500/20">
                    Flat Code: {currentRoomCode}
                  </span>
                )}
              </div>
              <h2 className="text-xl font-bold ledger-display text-slate-900 mt-1">{headerContent.title}</h2>
            </div>

            <div className="w-full md:w-auto grid grid-cols-2 md:flex items-center gap-1.5 bg-slate-900/5 p-1 rounded border" style={{ borderColor: 'var(--rule)' }}>
              <button
                onClick={() => setAnalysisType('personal')}
                className={`px-3 py-1 rounded text-xs font-bold transition-all flex items-center justify-center ${analysisType === 'personal' ? 'bg-[var(--ink)] text-[var(--card)] shadow-sm' : 'text-slate-600 hover:bg-slate-800/10'}`}
              >
                <Icons.User className="w-3.5 h-3.5 mr-1 inline" /> {currentUser?.name?.split(' ')[0] || 'My'}'s Personal
              </button>
              <button
                onClick={() => setAnalysisType('roommates')}
                className={`px-3 py-1 rounded text-xs font-bold transition-all flex items-center justify-center ${analysisType === 'roommates' ? 'bg-[var(--ink)] text-[var(--card)] shadow-sm' : 'text-slate-600 hover:bg-slate-800/10'}`}
              >
                <Icons.Group className="w-3.5 h-3.5 mr-1 inline" /> Roommates Split
              </button>
            </div>

            {analysisType === 'personal' && (
              <div className="flex items-center gap-2 bg-slate-900/5 p-1.5 rounded border text-xs" style={{ borderColor: 'var(--rule)' }}>
                <span className="font-semibold text-slate-500">{currentUser?.name?.split(' ')[0] || 'My'}'s Income:</span>
                <input
                  type="number"
                  className="w-16 bg-white border rounded px-1.5 py-0.5 text-center font-mono font-bold outline-none"
                  style={{ borderColor: 'var(--rule)' }}
                  value={salaryInput}
                  onChange={e => setSalaryInput(e.target.value)}
                />
                <button onClick={saveSalary} className="px-2 py-0.5 bg-[var(--ink)] text-[var(--card)] rounded text-[10px] font-bold">Save</button>
              </div>
            )}
          </div>

          <div style={{ display: (analysisType === 'roommates' && (roomMembershipStatus === 'none' || roomMembershipStatus === 'pending')) ? 'none' : 'block' }}>

          {/* ============================================================== */}
          {/* TAB 1: DASHBOARD OVERVIEW */}
          {/* ============================================================== */}
          {activeTab === 'dashboard' && (
            <div className="space-y-6">
              
              {/* Top rollup cards */}
              <div className={`grid grid-cols-2 ${analysisType === 'personal' ? 'lg:grid-cols-4' : 'lg:grid-cols-3'} gap-4`}>
                <div className="blur-card rounded p-4 flex flex-col justify-between">
                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Today Spent</span>
                  <span className="text-lg font-bold font-mono text-[var(--ink)] mt-1">
                    {analysisType === 'personal' 
                      ? fmt(totals.today) 
                      : fmt(totals.today / (roommates.length + 1))
                    }
                  </span>
                  <span className="text-[9px] text-slate-400 mt-0.5 font-mono">
                    {analysisType === 'personal' 
                      ? 'Personal expenses' 
                      : `Your share (Flat: ${fmt(totals.today)})`
                    }
                  </span>
                </div>
                <div className="blur-card rounded p-4 flex flex-col justify-between">
                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Weekly Pace</span>
                  <span className="text-lg font-bold font-mono text-[var(--ink)] mt-1">
                    {analysisType === 'personal' 
                      ? fmt(totals.week) 
                      : fmt(totals.week / (roommates.length + 1))
                    }
                  </span>
                  <span className="text-[9px] text-slate-400 mt-0.5 font-mono">
                    {analysisType === 'personal' 
                      ? 'Personal active week' 
                      : `Your share (Flat: ${fmt(totals.week)})`
                    }
                  </span>
                </div>
                <div className="blur-card rounded p-4 flex flex-col justify-between">
                  <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Monthly Total</span>
                  <span className="text-lg font-bold font-mono text-[var(--ink)] mt-1">
                    {analysisType === 'personal' 
                      ? fmt(totals.month) 
                      : fmt(totals.month / (roommates.length + 1))
                    }
                  </span>
                  <span className="text-[9px] text-slate-400 mt-0.5 font-mono">
                    {analysisType === 'personal' 
                      ? 'Personal billing month' 
                      : `Your share (Flat: ${fmt(totals.month)})`
                    }
                  </span>
                </div>
                {analysisType === 'personal' && (
                  <div className="blur-card rounded p-4 flex flex-col justify-between">
                    <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Yearly Rollup</span>
                    <span className="text-lg font-bold font-mono text-[var(--ink)] mt-1">
                      {fmt(totals.year)}
                    </span>
                    <span className="text-[9px] text-slate-400 mt-0.5 font-mono">
                      Personal calendar year
                    </span>
                  </div>
                )}
              </div>

              {/* Middle Section: Meter / Gauges */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* Salary Meter (Personal only) */}
                {/* Salary Meter (Personal) or Stats Card (Roommates) */}
                {analysisType === 'personal' ? (
                  <div className="blur-card rounded p-5 lg:col-span-1 flex flex-col justify-between">
                    <div>
                      <h3 className="font-bold text-xs uppercase tracking-wider text-slate-800 mb-1">Monthly Income Buffer</h3>
                      <p className="text-[10px] text-slate-500">Remaining personal salary after splits</p>
                    </div>

                    <div className="my-4 space-y-2">
                      <div className="w-full h-3 bg-slate-900/5 rounded-full overflow-hidden border" style={{ borderColor: 'var(--rule)' }}>
                        <div 
                          className="h-full transition-all duration-500"
                          style={{ 
                            width: `${Math.max(0, Math.min(100, (remainingSalary / salary) * 100))}%`,
                            background: remainingSalary > 10000 ? 'var(--positive)' : 'var(--stamp)'
                          }}
                        />
                      </div>
                      <div className="flex justify-between text-xs font-mono font-bold">
                        <span>Spent: {fmt(totals.month)}</span>
                        <span style={{ color: remainingSalary > 10000 ? 'var(--positive)' : 'var(--stamp)' }}>
                          Buffer: {fmt(remainingSalary)}
                        </span>
                      </div>
                    </div>

                    <div className="text-[10px] text-slate-500">
                      {currentUser?.name?.split(' ')[0] || 'Your'}'s personal target savings: 35% ({fmt(salary * 0.35)}).
                    </div>
                  </div>
                ) : (
                  <div className="blur-card rounded p-5 lg:col-span-1 flex flex-col justify-between">
                    <div>
                      <h3 className="font-bold text-xs uppercase tracking-wider text-slate-800 mb-1">Flat Activity & Stats</h3>
                      <p className="text-[10px] text-slate-500">Shared expense contributor highlights</p>
                    </div>
                    
                    <div className="space-y-2.5 my-3">
                      <div className="flex justify-between items-center p-2.5 bg-slate-900/5 border border-slate-200/60 hover:bg-slate-900/10 hover:shadow-sm rounded-lg transition-all duration-200">
                        <div className="flex items-center gap-2.5">
                          <span className="text-lg">🥇</span>
                          <div className="flex flex-col">
                            <span className="text-slate-800 font-bold text-xs">Highest Contributor</span>
                            <span className="text-[9px] text-slate-500 font-medium">Most rupees paid</span>
                          </div>
                        </div>
                        <div className="text-right">
                          <div className="font-bold text-emerald-800 text-xs">{roommateDues.highestContributor.name}</div>
                          <div className="text-[10px] text-slate-500 font-mono font-bold">{fmt(roommateDues.highestContributor.amount)}</div>
                        </div>
                      </div>
                      
                      <div className="flex justify-between items-center p-2.5 bg-slate-900/5 border border-slate-200/60 hover:bg-slate-900/10 hover:shadow-sm rounded-lg transition-all duration-200">
                        <div className="flex items-center gap-2.5">
                          <span className="text-lg">🥈</span>
                          <div className="flex flex-col">
                            <span className="text-slate-800 font-bold text-xs">Lowest Contributor</span>
                            <span className="text-[9px] text-slate-500 font-medium">Least rupees paid</span>
                          </div>
                        </div>
                        <div className="text-right">
                          <div className="font-bold text-amber-800 text-xs">{roommateDues.lowestContributor.name}</div>
                          <div className="text-[10px] text-slate-500 font-mono font-bold">{roommateDues.lowestContributor.amount === Infinity ? fmt(0) : fmt(roommateDues.lowestContributor.amount)}</div>
                        </div>
                      </div>
                      
                      <div className="flex justify-between items-center p-2.5 bg-slate-900/5 border border-slate-200/60 hover:bg-slate-900/10 hover:shadow-sm rounded-lg transition-all duration-200">
                        <div className="flex items-center gap-2.5">
                          <span className="text-lg">🛒</span>
                          <div className="flex flex-col">
                            <span className="text-slate-800 font-bold text-xs">Most Active Buyer</span>
                            <span className="text-[9px] text-slate-500 font-medium">Highest bill count</span>
                          </div>
                        </div>
                        <div className="text-right">
                          <div className="font-bold text-indigo-800 text-xs">{roommateDues.mostActiveBuyer.name}</div>
                          <div className="text-[10px] text-slate-500 font-mono font-bold">{roommateDues.mostActiveBuyer.count} bills</div>
                        </div>
                      </div>
                    </div>

                    <div className="text-[9px] text-slate-500 mt-1 uppercase font-bold">
                      Stats updated dynamically
                    </div>
                  </div>
                )}

                {/* Score Dial (Personal) or Dues summary (Roommates) */}
                {analysisType === 'personal' ? (
                  <div className="blur-card rounded p-5 flex items-center justify-between lg:col-span-1">
                    <div className="space-y-1.5">
                      <h3 className="font-bold text-xs uppercase tracking-wider text-slate-800">Bachelor Health</h3>
                      <div className="text-2xl font-extrabold ledger-display" style={{ color: healthScore > 75 ? 'var(--positive)' : healthScore > 50 ? '#D97706' : 'var(--stamp)' }}>
                        {healthScore}/100
                      </div>
                      <span className="text-[10px] font-bold uppercase border-2 px-1.5 py-0.2 rounded" style={{ borderColor: riskStatus.color, color: riskStatus.color }}>
                        {riskStatus.label}
                      </span>
                    </div>
                    <div className="w-16 h-16 relative flex items-center justify-center">
                      <svg className="w-full h-full transform -rotate-90" viewBox="0 0 36 36">
                        <circle cx="18" cy="18" r="16" fill="none" stroke="rgba(0,0,0,0.05)" strokeWidth="3" />
                        <circle 
                          cx="18" 
                          cy="18" 
                          r="16" 
                          fill="none" 
                          stroke={healthScore > 75 ? 'var(--positive)' : healthScore > 50 ? '#D97706' : 'var(--stamp)'}
                          strokeWidth="3" 
                          strokeDasharray="100 100" 
                          strokeDashoffset={100 - healthScore}
                        />
                      </svg>
                      <span className="absolute text-[10px] font-mono font-bold text-slate-600">{healthScore}%</span>
                    </div>
                  </div>
                ) : (
                  <div className="blur-card rounded p-5 flex flex-col justify-between lg:col-span-1">
                    <div>
                      <h3 className="font-bold text-xs uppercase tracking-wider text-slate-800 mb-1">Roommates Balance Status</h3>
                      <p className="text-[10px] text-slate-500">Member summaries (Paid vs equal share)</p>
                    </div>
                    
                    <div className="overflow-x-auto my-3">
                      <table className="w-full text-left text-[11px] font-mono">
                        <thead>
                          <tr className="border-b border-slate-900/10 text-[9px] uppercase tracking-wider text-slate-500 font-sans">
                            <th className="pb-1.5 font-bold">Member</th>
                            <th className="pb-1.5 text-right font-bold">Paid</th>
                            <th className="pb-1.5 text-right font-bold">Share</th>
                            <th className="pb-1.5 text-right font-bold">Status</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-900/5">
                          {roommateDues.memberSummaries.map((m) => {
                            const isMe = m.name === currentUser?.name;
                            let statusColor = 'var(--slate-600)';
                            let statusLabel = 'Settled';
                            let displayAmt = '';
                            
                            if (m.balance > 0) {
                              statusColor = 'var(--positive)';
                              statusLabel = 'Recv';
                              displayAmt = ` ₹${m.amountToReceive}`;
                            } else if (m.balance < 0) {
                              statusColor = 'var(--stamp)';
                              statusLabel = 'Pay';
                              displayAmt = ` ₹${m.amountToPay}`;
                            }

                            return (
                              <tr key={m.name} className="hover:bg-slate-900/5 transition-colors">
                                <td className="py-1.5 pr-1 font-sans text-slate-800 font-semibold truncate max-w-[80px]">
                                  {m.name} {isMe && '(You)'}
                                </td>
                                <td className="py-1.5 text-right text-slate-700">{fmt(m.paid)}</td>
                                <td className="py-1.5 text-right text-slate-500">{fmt(m.share)}</td>
                                <td className="py-1.5 text-right font-bold" style={{ color: statusColor }}>
                                  {statusLabel === 'Settled' ? 'Settled' : `${statusLabel}${displayAmt}`}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>

                    <div className="text-[9px] text-slate-500 font-bold uppercase mt-1">
                      Formula: Paid - Share = Balance
                    </div>
                  </div>
                )}

                {/* Flat Rent Split Summary (Roommates) or Swiggy Leakage Dial (Personal) */}
                <div className="blur-card rounded p-5 flex flex-col justify-between lg:col-span-1">
                  <div>
                    <h3 className="font-bold text-xs uppercase tracking-wider text-slate-800 mb-1">
                      {analysisType === 'personal' ? "Discretionary leakage pace" : "Flat rent & expenses"}
                    </h3>
                    <p className="text-[10px] text-slate-500">
                      {analysisType === 'personal' ? "Swiggy orders, gaming, retail spent" : "Rent, bills and shared splits proportion"}
                    </p>
                  </div>
                  
                  {analysisType === 'personal' ? (
                    <div className="flex justify-between items-baseline mt-3">
                      <div>
                        <span className="text-[9px] text-slate-400 uppercase block">Monthly total</span>
                        <span className="text-lg font-bold font-mono text-[var(--stamp)]">{fmt(discretionarySum)}</span>
                      </div>
                      <div className="text-right">
                        <span className="text-[9px] text-slate-400 uppercase block">Savings potential</span>
                        <span className="text-xs font-bold font-mono text-[var(--positive)]">+{fmt(discretionarySaved)}</span>
                      </div>
                    </div>
                  ) : (
                    (() => {
                      const loggedRentTotal = transactions.filter(t => t.is_shared && t.category === 'Rent' && t.date.slice(0, 7) === curMonthStr).reduce((sum, t) => sum + t.amount, 0);
                      const isLogged = loggedRentTotal > 0;
                      const rentVal = isLogged ? loggedRentTotal : parseFloat(flatRentInput || 12000);
                      
                      return (
                        <div className="flex flex-col gap-2 mt-3">
                          <div className="flex justify-between items-baseline border-b border-slate-900/5 pb-1.5">
                            <div>
                              <span className="text-[9px] text-slate-400 uppercase block">
                                Total Rent {!isLogged && <span className="text-[7.5px] text-slate-400 font-sans tracking-tight block font-normal">(Default setting — log to split)</span>}
                              </span>
                              <span className="text-sm font-bold font-mono text-slate-800">
                                {fmt(rentVal)}
                              </span>
                            </div>
                            <div className="text-right">
                              <span className="text-[9px] text-slate-400 uppercase block font-bold">Rent Share</span>
                              <span className="text-xs font-bold font-mono text-slate-600">
                                {fmt(rentVal / (roommates.length + 1))}
                              </span>
                            </div>
                          </div>
                          <div className="flex justify-between items-baseline mt-0.5">
                            <div>
                              <span className="text-[9px] text-slate-400 uppercase block font-semibold text-slate-500">Total Flat Spent</span>
                              <span className="text-sm font-bold font-mono text-slate-800">
                                {fmt(roommateDues.totalRoomExpense)}
                              </span>
                            </div>
                            <div className="text-right">
                              <span className="text-[9px] text-slate-400 uppercase block font-semibold text-slate-500">Your Share</span>
                              <span className="text-xs font-bold font-mono text-emerald-700">
                                {fmt(roommateDues.memberSummaries.find(m => m.name === currentUser?.name)?.share || 0)}
                              </span>
                            </div>
                          </div>

                          {/* Flat Rent Setting & Quick Log */}
                          <div className="mt-3 pt-3 border-t border-dashed border-slate-200">
                            <span className="text-[9px] text-slate-500 uppercase block font-bold tracking-wider">Configure Monthly Rent</span>
                            <div className="flex items-center gap-1.5 mt-1.5">
                              <div className="relative flex items-center">
                                <span className="absolute left-1.5 text-[9px] font-bold text-slate-400">₹</span>
                                <input
                                  type="number"
                                  className="w-20 bg-slate-100 border rounded pl-4 pr-1 py-0.5 text-center font-mono font-bold text-[10px] outline-none"
                                  style={{ borderColor: 'var(--rule)' }}
                                  value={flatRentInput}
                                  onChange={e => setFlatRentInput(e.target.value)}
                                  placeholder="12000"
                                  disabled={roomAdminId !== session?.user?.id}
                                />
                              </div>
                              {roomAdminId === session?.user?.id ? (
                                <>
                                  <button onClick={saveFlatRent} className="px-2 py-0.5 bg-[var(--ink)] text-[var(--card)] rounded text-[9px] font-bold">Save</button>
                                  <button onClick={quickLogRent} className="px-2 py-0.5 bg-[#5A6483] text-[#FBF8F0] hover:bg-[#4E5672] rounded text-[9px] font-bold flex items-center gap-0.5">
                                    Log rent split
                                  </button>
                                </>
                              ) : (
                                <span className="text-[9px] text-slate-400 italic">Admin only</span>
                              )}
                            </div>
                          </div>
                        </div>
                      );
                    })()
                  )}

                  <div className="text-[9px] text-slate-400 mt-2 font-mono">
                    {analysisType === 'personal' 
                      ? `Based on ${reductionPercent}% reduction slider settings`
                      : "Divided equally among accepted flat members"
                    }
                  </div>
                </div>

              </div>

              {/* Charts */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* Category Doughnut */}
                <div className="blur-card rounded p-5 flex flex-col justify-between lg:col-span-1">
                  <div className="text-[11px] font-bold ledger-display mb-3">
                    {analysisType === 'personal' ? "Personal Category Shares" : "Shared Flat Category Shares"}
                  </div>
                  
                  {totalMonthlySpend === 0 ? (
                    <div className="flex-1 flex items-center justify-center py-6 text-xs text-slate-500">No transactions recorded.</div>
                  ) : (
                    <div className="flex flex-col items-center gap-3">
                      <div className="relative w-24 h-24 flex-shrink-0 flex items-center justify-center">
                        <svg className="w-full h-full transform -rotate-90" viewBox="0 0 120 120">
                          {/* Background track circle */}
                          <circle cx="60" cy="60" r="50" fill="transparent" stroke="rgba(0, 0, 0, 0.04)" strokeWidth="10" />
                          {doughnutSegments.map((segment, idx) => (
                            <circle
                              key={idx}
                              cx="60"
                              cy="60"
                              r="50"
                              fill="transparent"
                              stroke={segment.color}
                              strokeWidth="10"
                              strokeDasharray={segment.strokeDasharray}
                              strokeDashoffset={segment.strokeDashoffset}
                              strokeLinecap="round"
                            />
                          ))}
                        </svg>
                        <div className="absolute flex flex-col items-center">
                          <span className="text-[8px] uppercase tracking-wider text-slate-500">Total</span>
                          <span className="text-xs font-bold font-mono text-[var(--ink)]">{fmt(totalMonthlySpend)}</span>
                        </div>
                      </div>

                      <div className="w-full text-xs space-y-1.5 max-h-36 overflow-y-auto pr-1">
                        {categoryBreakdown.slice(0, 5).map(cat => (
                          <div key={cat.name} className="flex items-center justify-between">
                            <div className="flex items-center gap-1.5 min-w-0">
                              <div className="w-5 h-5 rounded flex items-center justify-center flex-shrink-0" style={{ backgroundColor: `${cat.color}18`, color: cat.color }}>
                                <CategoryIcon category={cat.name} className="w-3 h-3" />
                              </div>
                              <span className="truncate text-slate-700">{cat.name}</span>
                            </div>
                            <span className="font-mono text-slate-800 font-bold">{fmt(cat.amount)}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* SVG Trend Line & ML Forecast */}
                <div className="blur-card rounded p-5 lg:col-span-2 flex flex-col justify-between">
                  <div className="flex items-center justify-between mb-3">
                    <div className="text-[11px] font-bold ledger-display">
                      Spend Trend & 3-Month ML Forecast
                    </div>
                    <div className="flex items-center gap-3 text-[9px] text-slate-500">
                      <span className="flex items-center gap-1"><span className="w-2 h-0.5 bg-[var(--ink)]" /> Historical</span>
                      <span className="flex items-center gap-1"><span className="w-2 h-0.5 border-t border-dashed border-[#7C3AED]" /> ML Forecast*</span>
                    </div>
                  </div>

                  <div className="w-full h-32 mt-1 relative">
                    {svgLineChartPath && (
                      <svg className="w-full h-full" viewBox="0 0 600 150" preserveAspectRatio="none">
                        <defs>
                          <linearGradient id="histGrad" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor="var(--ink)" stopOpacity="0.15" />
                            <stop offset="100%" stopColor="var(--ink)" stopOpacity="0.0" />
                          </linearGradient>
                          <linearGradient id="foreGrad" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor="#7C3AED" stopOpacity="0.15" />
                            <stop offset="100%" stopColor="#7C3AED" stopOpacity="0.0" />
                          </linearGradient>
                        </defs>

                        <line x1="40" y1="20" x2="560" y2="20" stroke="#e2e8f0" strokeWidth="0.5" strokeDasharray="2,2" />
                        <line x1="40" y1="75" x2="560" y2="75" stroke="#e2e8f0" strokeWidth="0.5" strokeDasharray="2,2" />
                        <line x1="40" y1="130" x2="560" y2="130" stroke="#e2e8f0" strokeWidth="0.5" strokeDasharray="2,2" />

                        <line 
                          x1="40" 
                          y1={150 - 20 - (salary / svgLineChartPath.maxVal) * (150 - 40)} 
                          x2="560" 
                          y2={150 - 20 - (salary / svgLineChartPath.maxVal) * (150 - 40)} 
                          stroke={`${currentTheme.stamp}35`} 
                          strokeWidth="1" 
                          strokeDasharray="4,2" 
                        />

                        {/* Smoothed linear gradients area fills */}
                        {svgLineChartPath.histAreaPath && (
                          <path d={svgLineChartPath.histAreaPath} fill="url(#histGrad)" />
                        )}
                        {svgLineChartPath.foreAreaPath && (
                          <path d={svgLineChartPath.foreAreaPath} fill="url(#foreGrad)" />
                        )}

                        <path d={svgLineChartPath.histPath} fill="none" stroke="var(--ink)" strokeWidth="3" strokeLinecap="round" />
                        <path d={svgLineChartPath.forePath} fill="none" stroke="#7C3AED" strokeWidth="3" strokeDasharray="4,4" strokeLinecap="round" />

                        {svgLineChartPath.coords.map((node, idx) => (
                          <circle
                            key={idx}
                            cx={node.x}
                            cy={node.y}
                            r={node.isForecast ? "3" : "4"}
                            fill={node.isForecast ? "var(--card)" : "var(--ink)"}
                            stroke={node.isForecast ? "#7C3AED" : "#ffffff"}
                            strokeWidth="1.5"
                          />
                        ))}
                      </svg>
                    )}
                  </div>

                  <div className="flex justify-between px-8 text-[9px] font-mono text-slate-500 mt-2">
                    {chartAndForecastData.map((pt, idx) => (
                      <span key={idx}>{pt.label}</span>
                    ))}
                  </div>
                </div>
              </div>

              {/* Alert Ribbon */}
              {analysisType === 'personal' && (
                <div className="flex flex-col md:flex-row md:items-center justify-between p-3.5 rounded border" style={{ borderColor: overToday > 0 ? `var(--stamp)30` : `var(--positive)30`, background: overToday > 0 ? `var(--stamp)05` : `var(--positive)05` }}>
                  <div className="flex items-center gap-2">
                    <span>{overToday > 0 ? <Icons.AlertWarning className="w-5 h-5 text-[var(--stamp)]" /> : <Icons.AlertSuccess className="w-5 h-5 text-[var(--positive)]" />}</span>
                    <div>
                      <h4 className="font-bold text-xs text-slate-900">
                        {overToday > 0 ? `Over budget today by ${fmt(overToday)}` : `Daily Budget: On Track`}
                      </h4>
                      <p className="text-[10px] text-slate-500">
                        {overToday > 0 
                          ? `You've exceeded ${currentUser?.name?.split(' ')[0] || 'your'}'s daily allowance. Cut back discretionary spent to balance.`
                          : `${currentUser?.name?.split(' ')[0] || 'You'} has ${fmt(-overToday)} left to spend today. Daily budget limit is ₹${Math.round(dailyBudget)}.`
                        }
                      </p>
                    </div>
                  </div>
                  <div className="text-right text-xs font-mono font-bold mt-1 md:mt-0" style={{ color: overToday > 0 ? 'var(--stamp)' : 'var(--positive)' }}>
                    Remaining Limit: {fmt(Math.max(0, -overToday))}
                  </div>
                </div>
              )}

              {/* Personal Ledger & Form on Dashboard (Personal Mode only) */}
              {analysisType === 'personal' && (
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mt-6">
                  {/* Detailed Form */}
                  <div className="blur-card rounded p-5 lg:col-span-1">
                    <h3 className="font-bold text-xs mb-3 text-slate-900 uppercase tracking-wider">Log Personal Expense</h3>
                    <form onSubmit={addTransaction} className="space-y-3.5">
                      <div className="flex flex-col">
                        <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Date</label>
                        <input
                          type="date"
                          className="ledger-input-box"
                          value={form.date}
                          onChange={e => setForm({ ...form, date: e.target.value })}
                        />
                      </div>
                      
                      <div className="grid grid-cols-2 gap-2">
                        <div className="flex flex-col">
                          <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Category</label>
                          <select
                            className="ledger-input-box cursor-pointer"
                            value={form.category}
                            onChange={e => setForm({ ...form, category: e.target.value, merchant: '' })}
                          >
                            {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>

                        <div className="flex flex-col">
                          <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Amount (₹)</label>
                          <input
                            type="number"
                            min="0"
                            placeholder="Amount"
                            className="ledger-input-box font-mono"
                            value={form.amount}
                            onChange={e => setForm({ ...form, amount: e.target.value })}
                          />
                        </div>
                      </div>

                      <div className="flex flex-col">
                        <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500 flex justify-between">
                          <span>Merchant / Vendor</span>
                          <span className="text-[9px] text-[var(--ink)] underline cursor-pointer" onClick={() => {
                            const list = MOCK_MERCHANTS[form.category] || ['Other'];
                            const r = list[Math.floor(Math.random() * list.length)];
                            setForm({ ...form, merchant: r });
                          }}>Suggest</span>
                        </label>
                        <input
                          type="text"
                          placeholder="e.g. Zomato, Coffee"
                          className="ledger-input-box"
                          value={form.merchant}
                          onChange={e => setForm({ ...form, merchant: e.target.value })}
                        />
                      </div>

                      <div className="flex flex-col">
                        <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Note</label>
                        <input
                          type="text"
                          placeholder="optional note"
                          className="ledger-input-box"
                          value={form.note}
                          onChange={e => setForm({ ...form, note: e.target.value })}
                        />
                      </div>

                      <button type="submit" onClick={() => setForm(f => ({ ...f, isShared: false }))} className="w-full btn-vintage-ink">
                        Add Personal Expense
                      </button>
                    </form>
                  </div>

                  {/* Personal Ledger List */}
                  <div className="blur-card rounded p-5 lg:col-span-2 flex flex-col justify-between">
                    <div>
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between pb-2 mb-4 border-b gap-2" style={{ borderColor: 'var(--rule)' }}>
                        <h3 className="font-bold text-xs text-slate-900 uppercase tracking-wider">
                          {currentUser?.name?.split(' ')[0] || 'My'}'s Personal Ledger
                        </h3>
                        {/* Export Dropdowns */}
                        <div className="flex gap-2 relative">
                          <div className="relative">
                            <button
                              onClick={() => { setExportMenuOpen(o => !o); setPdfMenuOpen(false); }}
                              className="btn-vintage-outline flex items-center gap-1 text-[10px]"
                            >
                              <Icons.Excel className="w-3.5 h-3.5 mr-1 inline" /> Export to Excel ▾
                            </button>
                            {exportMenuOpen && (
                              <div
                                className="absolute right-0 top-8 z-50 bg-white border rounded shadow-lg text-[11px] font-semibold text-slate-700 min-w-[140px] overflow-hidden"
                                style={{ borderColor: 'var(--rule)' }}
                              >
                                {[['today', '📅 Today'], ['week', '🗓 This Week (Sun–Sat)'], ['month', '📆 This Month'], ['year', '🗃 This Year'], ['all', '📋 All Time']].map(([val, label]) => (
                                  <button
                                    key={val}
                                    className="w-full text-left px-3 py-2 hover:bg-slate-100 border-b last:border-0 flex items-center gap-1"
                                    style={{ borderColor: '#e5e7eb' }}
                                    onClick={() => { exportToExcelCustom(val); setExportMenuOpen(false); }}
                                  >
                                    {label}
                                  </button>
                                ))}
                              </div>
                            )}
                          </div>

                          <div className="relative">
                            <button
                              onClick={() => { setPdfMenuOpen(o => !o); setExportMenuOpen(false); }}
                              className="btn-vintage-outline flex items-center gap-1 text-[10px]"
                            >
                              <Icons.PDF className="w-3.5 h-3.5 mr-1 inline" /> Export to PDF ▾
                            </button>
                            {pdfMenuOpen && (
                              <div
                                className="absolute right-0 top-8 z-50 bg-white border rounded shadow-lg text-[11px] font-semibold text-slate-700 min-w-[140px] overflow-hidden"
                                style={{ borderColor: 'var(--rule)' }}
                              >
                                {[['today', '📅 Today'], ['week', '🗓 This Week (Sun–Sat)'], ['month', '📆 This Month'], ['year', '🗃 This Year'], ['all', '📋 All Time']].map(([val, label]) => (
                                  <button
                                    key={val}
                                    className="w-full text-left px-3 py-2 hover:bg-slate-100 border-b last:border-0 flex items-center gap-1"
                                    style={{ borderColor: '#e5e7eb' }}
                                    onClick={() => { exportToPDFCustom(val); setPdfMenuOpen(false); }}
                                  >
                                    {label}
                                  </button>
                                ))}
                              </div>
                            )}
                          </div>
                        </div>
                      </div>

                      {/* Filters */}
                      <div className="flex flex-wrap gap-2 mb-4">
                        <div className="relative flex items-center flex-1 min-w-[120px]">
                          <input
                            type="text"
                            placeholder="Search..."
                            className="w-full bg-white border rounded pl-7 pr-2 py-1 text-[11px] outline-none text-slate-800"
                            style={{ borderColor: 'var(--rule)' }}
                            value={searchQuery}
                            onChange={e => setSearchQuery(e.target.value)}
                          />
                          <div className="absolute left-2.5"><Icons.Search className="w-3.5 h-3.5" /></div>
                        </div>
                        <select
                          className="bg-white border rounded px-2 py-1 text-[11px] outline-none cursor-pointer text-slate-800"
                          style={{ borderColor: 'var(--rule)' }}
                          value={filterCategory}
                          onChange={e => setFilterCategory(e.target.value)}
                        >
                          <option value="All">All Categories</option>
                          {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                        <div className="flex rounded border overflow-hidden text-[10px] font-bold" style={{ borderColor: 'var(--rule)' }}>
                          {[['all','All'],['today','Today'],['week','Week'],['month','Month'],['year','Year']].map(([val, lbl]) => (
                            <button
                              key={val}
                              onClick={() => setLedgerPeriod(val)}
                              className={`px-2 py-1 border-r last:border-0 transition ${
                                ledgerPeriod === val
                                  ? 'bg-[var(--ink)] text-[var(--card)]'
                                  : 'bg-white text-slate-600 hover:bg-slate-100'
                              }`}
                              style={{ borderColor: 'var(--rule)' }}
                            >
                              {lbl}
                            </button>
                          ))}
                        </div>
                      </div>

                      {/* Day-by-Day grouped transactions */}
                      {(() => {
                        const todayStrVal = isoDate(new Date());
                        const weekStartVal = new Date();
                        weekStartVal.setDate(weekStartVal.getDate() - weekStartVal.getDay());
                        const weekStartStrVal = isoDate(weekStartVal);
                        const monthStartStrVal = `${new Date().getFullYear()}-${String(new Date().getMonth()+1).padStart(2,'0')}-01`;
                        const yearStartStrVal = `${new Date().getFullYear()}-01-01`;

                        let baseTxs = processedTxs;

                        if (ledgerPeriod === 'today') baseTxs = baseTxs.filter(t => t.date === todayStrVal);
                        else if (ledgerPeriod === 'week') baseTxs = baseTxs.filter(t => t.date >= weekStartStrVal && t.date <= todayStrVal);
                        else if (ledgerPeriod === 'month') baseTxs = baseTxs.filter(t => t.date >= monthStartStrVal && t.date <= todayStrVal);
                        else if (ledgerPeriod === 'year') baseTxs = baseTxs.filter(t => t.date >= yearStartStrVal && t.date <= todayStrVal);

                        if (baseTxs.length === 0) return (
                          <div className="text-center py-8 text-xs text-slate-500">
                            📭 No personal expenses recorded.
                          </div>
                        );

                        const grouped = {};
                        baseTxs.forEach(tx => {
                          if (!grouped[tx.date]) grouped[tx.date] = [];
                          grouped[tx.date].push(tx);
                        });
                        const sortedDates = Object.keys(grouped).sort((a, b) => b.localeCompare(a));

                        return (
                          <div className="max-h-[350px] overflow-y-auto space-y-4 pr-1">
                            {sortedDates.map(date => {
                              const dayTxs = grouped[date];
                              const dayTotal = dayTxs.reduce((s, t) => s + Number(t.amount), 0);
                              const dateObj = new Date(date);
                              const isToday = date === todayStrVal;
                              const dayLabel = isToday ? 'Today' : dateObj.toLocaleDateString('en-IN', { weekday: 'short' });
                              const fullDate = dateObj.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });

                              return (
                                <div key={date}>
                                  <div className="flex items-center justify-between mb-1.5">
                                    <div className="flex items-center gap-2">
                                      <span className={`text-[10px] font-bold px-2 py-0.5 rounded ${
                                        isToday ? 'bg-[var(--ink)] text-[var(--card)]' : 'bg-slate-200 text-slate-700'
                                      }`}>{dayLabel}</span>
                                      <span className="text-[10px] text-slate-500 font-mono">{fullDate}</span>
                                    </div>
                                    <span className="text-[10px] font-bold font-mono text-[var(--stamp)]">
                                      Day total: {fmt(dayTotal)}
                                    </span>
                                  </div>
                                  <div className="space-y-1.5">
                                    {dayTxs.map(tx => (
                                      <div key={tx.id} className="p-2.5 rounded border hover:bg-slate-900/5 transition flex items-center justify-between gap-2" style={{ borderColor: 'var(--rule)', background: 'rgba(255,255,255,0.3)' }}>
                                        <div className="min-w-0 flex items-center gap-2.5">
                                          <div className="w-7 h-7 rounded-full flex items-center justify-center flex-shrink-0" style={{ backgroundColor: `${CATEGORY_COLORS[tx.category] || '#4B5563'}18`, color: CATEGORY_COLORS[tx.category] || '#4B5563' }}>
                                            <CategoryIcon category={tx.category} className="w-3.5 h-3.5" />
                                          </div>
                                          <div className="min-w-0">
                                            <div className="text-xs font-bold text-slate-900 flex items-center gap-1.5 flex-wrap">
                                              <span className="truncate text-slate-800">{tx.merchant}{tx.note ? ` — ${tx.note}` : ''}</span>
                                            </div>
                                            <div className="text-[9px] flex gap-1.5 items-center mt-0.5 text-slate-500">
                                              <span className="uppercase text-[7px] bg-[var(--paper)] px-1 py-0.5 rounded font-semibold text-slate-700">{tx.category}</span>
                                              <span>•</span>
                                              <span className="capitalize">{tx.source}</span>
                                            </div>
                                          </div>
                                        </div>
                                        <div className="flex items-center gap-3 flex-shrink-0">
                                          <span className="font-bold text-xs font-mono text-slate-900">{fmt(tx.amount)}</span>
                                          <button onClick={() => deleteTransaction(tx.id)} aria-label="Delete" className="text-red-600 hover:text-red-800 transition">
                                            <Icons.Trash className="w-3.5 h-3.5" />
                                          </button>
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        );
                      })()}
                    </div>
                  </div>
                </div>
              )}

              {/* Roommate Splits Ledger & Form on Dashboard (Roommates Mode only) */}
              {analysisType === 'roommates' && (
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mt-6">
                  {/* Detailed Form */}
                  <div className="blur-card rounded p-5 lg:col-span-1">
                    <h3 className="font-bold text-xs mb-3 text-slate-900 uppercase tracking-wider">Log Shared Roommate Expense</h3>
                    <form onSubmit={addTransaction} className="space-y-3.5">
                      <div className="flex flex-col">
                        <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Date</label>
                        <input
                          type="date"
                          className="ledger-input-box"
                          value={form.date}
                          onChange={e => setForm({ ...form, date: e.target.value })}
                        />
                      </div>
                      
                      <div className="grid grid-cols-2 gap-2">
                        <div className="flex flex-col">
                          <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Category</label>
                          <select
                            className="ledger-input-box cursor-pointer"
                            value={form.category}
                            onChange={e => setForm({ ...form, category: e.target.value, merchant: '' })}
                          >
                            {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>

                        <div className="flex flex-col">
                          <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Amount (₹)</label>
                          <input
                            type="number"
                            min="0"
                            placeholder="Amount"
                            className="ledger-input-box font-mono"
                            value={form.amount}
                            onChange={e => setForm({ ...form, amount: e.target.value })}
                          />
                        </div>
                      </div>

                      <div className="flex flex-col">
                        <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500 flex justify-between">
                          <span>Merchant / Vendor</span>
                          <span className="text-[9px] text-[var(--ink)] underline cursor-pointer" onClick={() => {
                            const list = MOCK_MERCHANTS[form.category] || ['Other'];
                            const r = list[Math.floor(Math.random() * list.length)];
                            setForm({ ...form, merchant: r });
                          }}>Suggest</span>
                        </label>
                        <input
                          type="text"
                          placeholder="e.g. Sri bought rice"
                          className="ledger-input-box"
                          value={form.merchant}
                          onChange={e => setForm({ ...form, merchant: e.target.value })}
                        />
                      </div>

                      <div className="flex flex-col">
                        <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Note</label>
                        <input
                          type="text"
                          placeholder="optional note"
                          className="ledger-input-box"
                          value={form.note}
                          onChange={e => setForm({ ...form, note: e.target.value })}
                        />
                      </div>

                      <button type="submit" onClick={() => setForm(f => ({ ...f, isShared: true }))} className="w-full btn-vintage-ink">
                        Add Shared Split
                      </button>
                    </form>
                  </div>

                  {/* Shared Ledger List */}
                  <div className="blur-card rounded p-5 lg:col-span-2 flex flex-col justify-between">
                    <div>
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between pb-2 mb-4 border-b gap-2" style={{ borderColor: 'var(--rule)' }}>
                        <h3 className="font-bold text-xs text-slate-900 uppercase tracking-wider">
                          Shared Roommates Ledger
                        </h3>
                        {/* Export Dropdowns */}
                        <div className="flex gap-2 relative">
                          <div className="relative">
                            <button
                              onClick={() => { setExportMenuOpen(o => !o); setPdfMenuOpen(false); }}
                              className="btn-vintage-outline flex items-center gap-1 text-[10px]"
                            >
                              <Icons.Excel className="w-3.5 h-3.5 mr-1 inline" /> Export to Excel ▾
                            </button>
                            {exportMenuOpen && (
                              <div
                                className="absolute right-0 top-8 z-50 bg-white border rounded shadow-lg text-[11px] font-semibold text-slate-700 min-w-[140px] overflow-hidden"
                                style={{ borderColor: 'var(--rule)' }}
                              >
                                {[['today', '📅 Today'], ['week', '🗓 This Week (Sun–Sat)'], ['month', '📆 This Month'], ['year', '🗃 This Year'], ['all', '📋 All Time']].map(([val, label]) => (
                                  <button
                                    key={val}
                                    className="w-full text-left px-3 py-2 hover:bg-slate-100 border-b last:border-0 flex items-center gap-1"
                                    style={{ borderColor: '#e5e7eb' }}
                                    onClick={() => { exportToExcelCustom(val); setExportMenuOpen(false); }}
                                  >
                                    {label}
                                  </button>
                                ))}
                              </div>
                            )}
                          </div>

                          <div className="relative">
                            <button
                              onClick={() => { setPdfMenuOpen(o => !o); setExportMenuOpen(false); }}
                              className="btn-vintage-outline flex items-center gap-1 text-[10px]"
                            >
                              <Icons.PDF className="w-3.5 h-3.5 mr-1 inline" /> Export to PDF ▾
                            </button>
                            {pdfMenuOpen && (
                              <div
                                className="absolute right-0 top-8 z-50 bg-white border rounded shadow-lg text-[11px] font-semibold text-slate-700 min-w-[140px] overflow-hidden"
                                style={{ borderColor: 'var(--rule)' }}
                              >
                                {[['today', '📅 Today'], ['week', '🗓 This Week (Sun–Sat)'], ['month', '📆 This Month'], ['year', '🗃 This Year'], ['all', '📋 All Time']].map(([val, label]) => (
                                  <button
                                    key={val}
                                    className="w-full text-left px-3 py-2 hover:bg-slate-100 border-b last:border-0 flex items-center gap-1"
                                    style={{ borderColor: '#e5e7eb' }}
                                    onClick={() => { exportToPDFCustom(val); setPdfMenuOpen(false); }}
                                  >
                                    {label}
                                  </button>
                                ))}
                              </div>
                            )}
                          </div>
                        </div>
                      </div>

                      {/* Filters */}
                      <div className="flex flex-wrap gap-2 mb-4">
                        <div className="relative flex items-center flex-1 min-w-[120px]">
                          <input
                            type="text"
                            placeholder="Search..."
                            className="w-full bg-white border rounded pl-7 pr-2 py-1 text-[11px] outline-none text-slate-800"
                            style={{ borderColor: 'var(--rule)' }}
                            value={searchQuery}
                            onChange={e => setSearchQuery(e.target.value)}
                          />
                          <div className="absolute left-2.5"><Icons.Search className="w-3.5 h-3.5" /></div>
                        </div>
                        <select
                          className="bg-white border rounded px-2 py-1 text-[11px] outline-none cursor-pointer text-slate-800"
                          style={{ borderColor: 'var(--rule)' }}
                          value={filterCategory}
                          onChange={e => setFilterCategory(e.target.value)}
                        >
                          <option value="All">All Categories</option>
                          {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                        <div className="flex rounded border overflow-hidden text-[10px] font-bold" style={{ borderColor: 'var(--rule)' }}>
                          {[['all','All'],['today','Today'],['week','Week'],['month','Month'],['year','Year']].map(([val, lbl]) => (
                            <button
                              key={val}
                              onClick={() => setLedgerPeriod(val)}
                              className={`px-2 py-1 border-r last:border-0 transition ${
                                ledgerPeriod === val
                                  ? 'bg-[var(--ink)] text-[var(--card)]'
                                  : 'bg-white text-slate-600 hover:bg-slate-100'
                              }`}
                              style={{ borderColor: 'var(--rule)' }}
                            >
                              {lbl}
                            </button>
                          ))}
                        </div>
                      </div>

                      {/* Day-by-Day grouped transactions */}
                      {(() => {
                        const todayStrVal = isoDate(new Date());
                        const weekStartVal = new Date();
                        weekStartVal.setDate(weekStartVal.getDate() - weekStartVal.getDay());
                        const weekStartStrVal = isoDate(weekStartVal);
                        const monthStartStrVal = `${new Date().getFullYear()}-${String(new Date().getMonth()+1).padStart(2,'0')}-01`;
                        const yearStartStrVal = `${new Date().getFullYear()}-01-01`;

                        let baseTxs = processedTxs;

                        if (ledgerPeriod === 'today') baseTxs = baseTxs.filter(t => t.date === todayStrVal);
                        else if (ledgerPeriod === 'week') baseTxs = baseTxs.filter(t => t.date >= weekStartStrVal && t.date <= todayStrVal);
                        else if (ledgerPeriod === 'month') baseTxs = baseTxs.filter(t => t.date >= monthStartStrVal && t.date <= todayStrVal);
                        else if (ledgerPeriod === 'year') baseTxs = baseTxs.filter(t => t.date >= yearStartStrVal && t.date <= todayStrVal);

                        if (baseTxs.length === 0) return (
                          <div className="text-center py-8 text-xs text-slate-500">
                            📭 No shared expenses recorded.
                          </div>
                        );

                        const grouped = {};
                        baseTxs.forEach(tx => {
                          if (!grouped[tx.date]) grouped[tx.date] = [];
                          grouped[tx.date].push(tx);
                        });
                        const sortedDates = Object.keys(grouped).sort((a, b) => b.localeCompare(a));

                        return (
                          <div className="max-h-[350px] overflow-y-auto space-y-4 pr-1">
                            {sortedDates.map(date => {
                              const dayTxs = grouped[date];
                              const dayTotal = dayTxs.reduce((s, t) => s + Number(t.amount), 0);
                              const dateObj = new Date(date);
                              const isToday = date === todayStrVal;
                              const dayLabel = isToday ? 'Today' : dateObj.toLocaleDateString('en-IN', { weekday: 'short' });
                              const fullDate = dateObj.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });

                              return (
                                <div key={date}>
                                  <div className="flex items-center justify-between mb-1.5">
                                    <div className="flex items-center gap-2">
                                      <span className={`text-[10px] font-bold px-2 py-0.5 rounded ${
                                        isToday ? 'bg-[var(--ink)] text-[var(--card)]' : 'bg-slate-200 text-slate-700'
                                      }`}>{dayLabel}</span>
                                      <span className="text-[10px] text-slate-500 font-mono">{fullDate}</span>
                                    </div>
                                    <span className="text-[10px] font-bold font-mono text-[var(--stamp)]">
                                      Day total: {fmt(dayTotal)}
                                    </span>
                                  </div>
                                  <div className="space-y-1.5">
                                    {dayTxs.map(tx => (
                                      <div key={tx.id} className="p-3 rounded border hover:bg-slate-900/5 transition flex flex-col gap-2" style={{ borderColor: 'var(--rule)', background: 'rgba(255,255,255,0.3)' }}>
                                        <div className="flex items-center justify-between gap-2">
                                          <div className="min-w-0 flex items-center gap-2.5">
                                            <div className="w-7 h-7 rounded-full flex items-center justify-center flex-shrink-0" style={{ backgroundColor: `${CATEGORY_COLORS[tx.category] || '#4B5563'}18`, color: CATEGORY_COLORS[tx.category] || '#4B5563' }}>
                                              <CategoryIcon category={tx.category} className="w-3.5 h-3.5" />
                                            </div>
                                            <div className="min-w-0">
                                              <div className="text-xs font-bold text-slate-900 flex items-center gap-1.5 flex-wrap">
                                                <span className="px-1.5 py-0.5 rounded bg-[var(--ink)] text-[var(--card)] text-[9px] uppercase font-bold tracking-wider flex-shrink-0">
                                                  {tx.logged_by} paid
                                                </span>
                                                <span className="truncate text-slate-800">{tx.merchant}{tx.note ? ` — ${tx.note}` : ''}</span>
                                              </div>
                                              <div className="text-[9px] flex gap-1.5 items-center mt-0.5 text-slate-500">
                                                <span className="uppercase text-[7px] bg-[var(--paper)] px-1 py-0.5 rounded font-semibold text-slate-700">{tx.category}</span>
                                                <span>•</span>
                                                <span className="capitalize">{tx.source}</span>
                                                <span className="text-emerald-700 font-bold">👥 Shared</span>
                                              </div>
                                            </div>
                                          </div>
                                          <div className="flex items-center gap-3 flex-shrink-0">
                                            <span className="font-bold text-xs font-mono text-slate-900">{fmt(tx.amount)}</span>
                                            <button onClick={() => deleteTransaction(tx.id.replace('-split', ''))} aria-label="Delete" className="text-red-600 hover:text-red-800 transition">
                                              <Icons.Trash className="w-3.5 h-3.5" />
                                            </button>
                                          </div>
                                        </div>

                                        <div className="mt-1 pt-2 border-t border-dashed border-slate-200 text-[10px] text-slate-600">
                                          <div className="flex justify-between font-bold text-slate-700 mb-1 bg-slate-100/50 p-1.5 rounded">
                                            <span>Split count: {roommates.length + 1} people</span>
                                            <span className="text-emerald-700">₹{Math.round(tx.amount / (roommates.length + 1))} / head</span>
                                          </div>
                                          <div className="grid grid-cols-2 sm:grid-cols-3 gap-1.5 font-mono text-[9px] mt-1 bg-slate-900/5 p-1.5 rounded">
                                            <div className="flex justify-between border-b pb-0.5" style={{ borderColor: 'var(--rule)' }}>
                                              <span className="text-slate-500 truncate">{currentUser?.name === tx.logged_by ? `${currentUser?.name} (Payer)` : currentUser?.name}:</span>
                                              <span className="font-bold text-slate-800">₹{Math.round(tx.amount / (roommates.length + 1))}</span>
                                            </div>
                                            {roommates.map(r => (
                                              <div key={r.id} className="flex justify-between border-b pb-0.5" style={{ borderColor: 'var(--rule)' }}>
                                                <span className="text-slate-500 truncate">{r.name === tx.logged_by ? `${r.name} (Payer)` : r.name}:</span>
                                                <span className="font-bold text-slate-800">₹{Math.round(tx.amount / (roommates.length + 1))}</span>
                                              </div>
                                            ))}
                                          </div>
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        );
                      })()}
                    </div>
                  </div>
                </div>
              )}

            </div>
          )}

          {/* ============================================================== */}
          {/* TAB 2: TRANSACTIONS & EXCEL PIPELINE */}
          {/* ============================================================== */}
          {activeTab === 'transactions' && (
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              
              {/* Detailed Form */}
              <div className="blur-card rounded p-5 lg:col-span-1">
                <h3 className="font-bold text-xs mb-3 text-slate-900 uppercase tracking-wider">Log Ledger Expense</h3>
                <form onSubmit={addTransaction} className="space-y-3.5">
                  <div className="flex flex-col">
                    <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Date</label>
                    <input
                      type="date"
                      className="ledger-input-box"
                      value={form.date}
                      onChange={e => setForm({ ...form, date: e.target.value })}
                    />
                  </div>
                  
                  <div className="grid grid-cols-2 gap-2">
                    <div className="flex flex-col">
                      <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Category</label>
                      <select
                        className="ledger-input-box cursor-pointer"
                        value={form.category}
                        onChange={e => setForm({ ...form, category: e.target.value, merchant: '' })}
                      >
                        {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <div className="flex flex-col">
                      <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Amount (₹)</label>
                      <input
                        type="number"
                        min="0"
                        placeholder="Amount"
                        className="ledger-input-box font-mono"
                        value={form.amount}
                        onChange={e => setForm({ ...form, amount: e.target.value })}
                      />
                    </div>
                  </div>

                  <div className="flex flex-col">
                    <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500 flex justify-between">
                      <span>Merchant / Vendor</span>
                      <span className="text-[9px] text-[var(--ink)] underline cursor-pointer" onClick={() => {
                        const list = MOCK_MERCHANTS[form.category] || ['Other'];
                        const r = list[Math.floor(Math.random() * list.length)];
                        setForm({ ...form, merchant: r });
                      }}>Suggest</span>
                    </label>
                    <input
                      type="text"
                      placeholder="e.g. Uber, Zomato, Excitel WiFi"
                      className="ledger-input-box"
                      value={form.merchant}
                      onChange={e => setForm({ ...form, merchant: e.target.value })}
                    />
                  </div>

                  <div className="flex flex-col">
                    <label className="text-[10px] uppercase font-bold tracking-wider mb-1 text-slate-500">Note</label>
                    <input
                      type="text"
                      placeholder="optional note"
                      className="ledger-input-box"
                      value={form.note}
                      onChange={e => setForm({ ...form, note: e.target.value })}
                    />
                  </div>

                  {/* Only show roommate split checkbox in roommates mode */}
                  {analysisType === 'roommates' && roomMembershipStatus === 'accepted' && (
                    <div className="flex items-center gap-2 p-2 bg-slate-900/5 border rounded" style={{ borderColor: 'var(--rule)' }}>
                      <input
                        type="checkbox"
                        id="sharedCheckDetailed"
                        className="cursor-pointer"
                        checked={form.isShared}
                        onChange={e => setForm({ ...form, isShared: e.target.checked })}
                      />
                      <label htmlFor="sharedCheckDetailed" className="text-xs font-bold text-slate-800 cursor-pointer select-none">
                        <Icons.Group className="w-3.5 h-3.5 mr-1 inline text-[var(--ink-soft)]" /> Split equally with roommates?
                      </label>
                    </div>
                  )}

                  <button type="submit" className="w-full btn-vintage-ink">
                    Add Transaction
                  </button>
                </form>
              </div>

              {/* Transactions table */}
              <div className="blur-card rounded p-5 lg:col-span-2 flex flex-col justify-between">
                <div>
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between pb-2 mb-4 border-b gap-2" style={{ borderColor: 'var(--rule)' }}>
                    <h3 className="font-bold text-xs text-slate-900 uppercase tracking-wider">
                      {analysisType === 'personal' ? `${currentUser?.name?.split(' ')[0] || 'My'}'s Personal Ledger` : "Shared Roommates Ledger"}
                    </h3>
                    {/* Export Dropdowns */}
                    <div className="flex gap-2 relative">
                      <div className="relative">
                        <button
                          onClick={() => { setExportMenuOpen(o => !o); setPdfMenuOpen(false); }}
                          className="btn-vintage-outline flex items-center gap-1 text-[10px]"
                        >
                          <Icons.Excel className="w-3.5 h-3.5 mr-1 inline" /> Export to Excel ▾
                        </button>
                        {exportMenuOpen && (
                          <div
                            className="absolute right-0 top-8 z-50 bg-white border rounded shadow-lg text-[11px] font-semibold text-slate-700 min-w-[140px] overflow-hidden"
                            style={{ borderColor: 'var(--rule)' }}
                          >
                            {[['today', '📅 Today'], ['week', '🗓 This Week (Sun–Sat)'], ['month', '📆 This Month'], ['year', '🗃 This Year'], ['all', '📋 All Time']].map(([val, label]) => (
                              <button
                                key={val}
                                className="w-full text-left px-3 py-2 hover:bg-slate-100 border-b last:border-0 flex items-center gap-1"
                                style={{ borderColor: '#e5e7eb' }}
                                onClick={() => { exportToExcelCustom(val); setExportMenuOpen(false); }}
                              >
                                {label}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>

                      <div className="relative">
                        <button
                          onClick={() => { setPdfMenuOpen(o => !o); setExportMenuOpen(false); }}
                          className="btn-vintage-outline flex items-center gap-1 text-[10px]"
                        >
                          <Icons.PDF className="w-3.5 h-3.5 mr-1 inline" /> Export to PDF ▾
                        </button>
                        {pdfMenuOpen && (
                          <div
                            className="absolute right-0 top-8 z-50 bg-white border rounded shadow-lg text-[11px] font-semibold text-slate-700 min-w-[140px] overflow-hidden"
                            style={{ borderColor: 'var(--rule)' }}
                          >
                            {[['today', '📅 Today'], ['week', '🗓 This Week (Sun–Sat)'], ['month', '📆 This Month'], ['year', '🗃 This Year'], ['all', '📋 All Time']].map(([val, label]) => (
                              <button
                                key={val}
                                className="w-full text-left px-3 py-2 hover:bg-slate-100 border-b last:border-0 flex items-center gap-1"
                                style={{ borderColor: '#e5e7eb' }}
                                onClick={() => { exportToPDFCustom(val); setPdfMenuOpen(false); }}
                              >
                                {label}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Filters */}
                  <div className="flex flex-wrap gap-2 mb-4">
                    <div className="relative flex items-center flex-1 min-w-[120px]">
                      <input
                        type="text"
                        placeholder="Search..."
                        className="w-full bg-white border rounded pl-7 pr-2 py-1 text-[11px] outline-none text-slate-800"
                        style={{ borderColor: 'var(--rule)' }}
                        value={searchQuery}
                        onChange={e => setSearchQuery(e.target.value)}
                      />
                      <div className="absolute left-2.5"><Icons.Search className="w-3.5 h-3.5" /></div>
                    </div>
                    <select
                      className="bg-white border rounded px-2 py-1 text-[11px] outline-none cursor-pointer text-slate-800"
                      style={{ borderColor: 'var(--rule)' }}
                      value={filterCategory}
                      onChange={e => setFilterCategory(e.target.value)}
                    >
                      <option value="All">All Categories</option>
                      {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                    {/* Period filter */}
                    <div className="flex rounded border overflow-hidden text-[10px] font-bold" style={{ borderColor: 'var(--rule)' }}>
                      {[['all','All'],['today','Today'],['week','Week'],['month','Month'],['year','Year']].map(([val, lbl]) => (
                        <button
                          key={val}
                          onClick={() => setLedgerPeriod(val)}
                          className={`px-2 py-1 border-r last:border-0 transition ${
                            ledgerPeriod === val
                              ? 'bg-[var(--ink)] text-[var(--card)]'
                              : 'bg-white text-slate-600 hover:bg-slate-100'
                          }`}
                          style={{ borderColor: 'var(--rule)' }}
                        >
                          {lbl}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Day-by-Day grouped transactions */}
                  {(() => {
                    const today = new Date();
                    const todayStr = isoDate(today);
                    const weekStart = new Date(today);
                    weekStart.setDate(today.getDate() - today.getDay());
                    const weekStartStr = isoDate(weekStart);
                    const monthStartStr = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-01`;
                    const yearStartStr = `${today.getFullYear()}-01-01`;

                    // Filter by personal only
                    let baseTxs = analysisType === 'personal'
                      ? processedTxs.filter(t => !t.is_shared)
                      : processedTxs;

                    // Apply period filter
                    if (ledgerPeriod === 'today') baseTxs = baseTxs.filter(t => t.date === todayStr);
                    else if (ledgerPeriod === 'week') baseTxs = baseTxs.filter(t => t.date >= weekStartStr && t.date <= todayStr);
                    else if (ledgerPeriod === 'month') baseTxs = baseTxs.filter(t => t.date >= monthStartStr && t.date <= todayStr);
                    else if (ledgerPeriod === 'year') baseTxs = baseTxs.filter(t => t.date >= yearStartStr && t.date <= todayStr);

                    if (baseTxs.length === 0) return (
                      <div className="text-center py-8 text-xs text-slate-500">
                        {ledgerPeriod === 'today' ? '📭 No expenses logged today yet.' :
                         ledgerPeriod === 'week' ? '📭 No expenses this week.' :
                         ledgerPeriod === 'month' ? '📭 No expenses this month.' :
                         '📭 No transactions recorded.'}
                      </div>
                    );

                    // Group by date
                    const grouped = {};
                    baseTxs.forEach(tx => {
                      if (!grouped[tx.date]) grouped[tx.date] = [];
                      grouped[tx.date].push(tx);
                    });
                    const sortedDates = Object.keys(grouped).sort((a, b) => b.localeCompare(a));

                    return (
                      <div className="max-h-[400px] overflow-y-auto space-y-4 pr-1">
                        {sortedDates.map(date => {
                          const dayTxs = grouped[date];
                          const dayTotal = dayTxs.reduce((s, t) => s + Number(t.amount), 0);
                          const dateObj = new Date(date);
                          const isToday = date === todayStr;
                          const dayLabel = isToday ? 'Today' : dateObj.toLocaleDateString('en-IN', { weekday: 'short' });
                          const fullDate = dateObj.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });

                          return (
                            <div key={date}>
                              {/* Day header */}
                              <div className="flex items-center justify-between mb-1.5">
                                <div className="flex items-center gap-2">
                                  <span className={`text-[10px] font-bold px-2 py-0.5 rounded ${
                                    isToday ? 'bg-[var(--ink)] text-[var(--card)]' : 'bg-slate-200 text-slate-700'
                                  }`}>{dayLabel}</span>
                                  <span className="text-[10px] text-slate-500 font-mono">{fullDate}</span>
                                </div>
                                <span className="text-[10px] font-bold font-mono text-[var(--stamp)]">
                                  Day total: {fmt(dayTotal)}
                                </span>
                              </div>
                              {/* Day transactions */}
                              <div className="space-y-1.5">
                                {dayTxs.map(tx => (
                                  <div key={tx.id} className="p-3 rounded border hover:bg-slate-900/5 transition flex flex-col gap-2" style={{ borderColor: 'var(--rule)', background: 'rgba(255,255,255,0.3)' }}>
                                    {/* Top Row: Icon, Merchant, Amount, Delete */}
                                    <div className="flex items-center justify-between gap-2">
                                      <div className="min-w-0 flex items-center gap-2.5">
                                        <div className="w-7 h-7 rounded-full flex items-center justify-center flex-shrink-0" style={{ backgroundColor: `${CATEGORY_COLORS[tx.category] || '#4B5563'}18`, color: CATEGORY_COLORS[tx.category] || '#4B5563' }}>
                                          <CategoryIcon category={tx.category} className="w-3.5 h-3.5" />
                                        </div>
                                        <div className="min-w-0">
                                          <div className="text-xs font-bold text-slate-900 flex items-center gap-1.5 flex-wrap">
                                            {tx.is_shared && (
                                              <span className="px-1.5 py-0.5 rounded bg-[var(--ink)] text-[var(--card)] text-[9px] uppercase font-bold tracking-wider flex-shrink-0">
                                                {tx.logged_by} paid
                                              </span>
                                            )}
                                            <span className="truncate text-slate-800">{tx.merchant}{tx.note ? ` — ${tx.note}` : ''}</span>
                                          </div>
                                          <div className="text-[9px] flex gap-1.5 items-center mt-0.5 text-slate-500">
                                            <span className="uppercase text-[7px] bg-[var(--paper)] px-1 py-0.5 rounded font-semibold text-slate-700">{tx.category}</span>
                                            <span>•</span>
                                            <span className="capitalize">{tx.source}</span>
                                            {tx.is_shared && <span className="text-emerald-700 font-bold">👥 Shared</span>}
                                          </div>
                                        </div>
                                      </div>
                                      <div className="flex items-center gap-3 flex-shrink-0">
                                        <span className="font-bold text-xs font-mono text-slate-900">{fmt(tx.amount)}</span>
                                        <button onClick={() => deleteTransaction(tx.id.replace('-split', ''))} aria-label="Delete" className="text-red-600 hover:text-red-800 transition">
                                          <Icons.Trash className="w-3.5 h-3.5" />
                                        </button>
                                      </div>
                                    </div>

                                    {/* Bottom Row: Roommate Splits Details (only visible in roommates analysis mode) */}
                                    {analysisType === 'roommates' && tx.is_shared && (
                                      <div className="mt-1 pt-2 border-t border-dashed border-slate-200 text-[10px] text-slate-600">
                                        <div className="flex justify-between font-bold text-slate-700 mb-1 bg-slate-100/50 p-1.5 rounded">
                                          <span>Split count: {roommates.length + 1} people</span>
                                          <span className="text-emerald-700">₹{Math.round(tx.amount / (roommates.length + 1))} / head</span>
                                        </div>
                                        <div className="grid grid-cols-2 sm:grid-cols-3 gap-1.5 font-mono text-[9px] mt-1 bg-slate-900/5 p-1.5 rounded">
                                          <div className="flex justify-between border-b pb-0.5" style={{ borderColor: 'var(--rule)' }}>
                                            <span className="text-slate-500 truncate">{currentUser?.name === tx.logged_by ? `${currentUser?.name} (Payer)` : currentUser?.name}:</span>
                                            <span className="font-bold text-slate-800">₹{Math.round(tx.amount / (roommates.length + 1))}</span>
                                          </div>
                                          {roommates.map(r => (
                                            <div key={r.id} className="flex justify-between border-b pb-0.5" style={{ borderColor: 'var(--rule)' }}>
                                              <span className="text-slate-500 truncate">{r.name === tx.logged_by ? `${r.name} (Payer)` : r.name}:</span>
                                              <span className="font-bold text-slate-800">₹{Math.round(tx.amount / (roommates.length + 1))}</span>
                                            </div>
                                          ))}
                                        </div>
                                      </div>
                                    )}
                                  </div>
                                ))}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    );
                  })()}
                </div>

                <div className="mt-4 pt-3 border-t text-right text-[11px] text-slate-500 flex justify-between" style={{ borderColor: 'var(--rule)' }}>
                  <span>Period: <strong>{ledgerPeriod === 'all' ? 'All Time' : ledgerPeriod.charAt(0).toUpperCase() + ledgerPeriod.slice(1)}</strong></span>
                </div>
              </div>

            </div>
          )}

          {/* ============================================================== */}
          {/* TAB 3: RULE-BASED WASTE FINDER */}
          {/* ============================================================== */}
          {activeTab === 'waste' && (
            <div className="space-y-6">
              
              <div>
                <h3 className="font-bold text-sm mb-3 text-slate-900 uppercase tracking-wider flex items-center gap-1.5">
                  <Icons.Search className="w-4 h-4 mr-1 inline" /> Daily Leakage Finder <span className="text-[9px] px-1.5 py-0.2 rounded bg-red-600/10 text-red-800 font-bold border border-red-500/20">Personal ({currentUser?.name?.split(' ')[0] || 'Me'})</span>
                </h3>
                {leakageItems.length === 0 ? (
                  <div className="blur-card rounded p-6 text-center text-xs text-slate-500">
                    No recurring leakage pattern found. Add 5+ small tea, coffee, or snacks purchases under ₹150 to trigger this detector.
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {leakageItems.map((item, idx) => (
                      <div key={idx} className="blur-card rounded p-5 flex flex-col justify-between border-l-4 border-l-[var(--stamp)]">
                        <div>
                          <div className="flex justify-between items-baseline">
                            <h4 className="font-bold text-xs text-slate-900">{item.merchant} ({item.category})</h4>
                            <span className="text-[9px] font-mono font-bold text-[var(--stamp)]">{item.count} logs / month</span>
                          </div>
                          <p className="text-xs mt-2 text-slate-600">
                            Bachelor small daily charges sum up quietly {currentUser?.name?.split(' ')[0] || 'friend'}. You spent <span className="font-bold text-slate-900">{fmt(item.monthlyCost)}</span> last month on this item.
                          </p>
                        </div>
                        <div className="grid grid-cols-2 gap-2 mt-4 pt-3 border-t border-slate-200 text-center">
                          <div className="rounded bg-slate-900/5 p-2">
                            <span className="text-[9px] uppercase block text-slate-500">Yearly leakage</span>
                            <span className="text-sm font-bold font-mono text-[var(--stamp)]">{fmt(item.yearlyCost)}</span>
                          </div>
                          <div className="rounded bg-emerald-600/10 p-2">
                            <span className="text-[9px] uppercase block text-slate-600">If invested (5y @ 12%)</span>
                            <span className="text-sm font-bold font-mono text-[var(--positive)]">{fmt(item.projectedWealth)}</span>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div>
                <h3 className="font-bold text-sm mb-3 text-slate-900 uppercase tracking-wider flex items-center gap-1.5">
                  <Icons.Ingestion className="w-4 h-4 mr-1 inline" /> Subscription Detector & Usage Scanner
                </h3>
                {subscriptions.length === 0 ? (
                  <div className="blur-card rounded p-6 text-center text-xs text-slate-500">
                    No active subscriptions detected.
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    {subscriptions.map((sub, idx) => (
                      <div key={idx} className={`blur-card rounded p-5 flex flex-col justify-between border-t-2 ${sub.isLowUsage ? 'border-t-[var(--stamp)]' : 'border-t-[var(--ink)]'}`}>
                        <div>
                          <div className="flex justify-between items-baseline mb-2">
                            <h4 className="font-bold text-xs text-slate-900">{sub.name}</h4>
                            <span className="text-[9px] font-mono px-1.5 py-0.2 rounded bg-[var(--paper)] text-slate-800 font-bold">{fmt(sub.amount)} / mo</span>
                          </div>
                          <p className="text-[11px] text-slate-500">
                            Detected at interval of ~{sub.avgGap} days. 
                          </p>

                          <div className="mt-3 bg-white p-2.5 rounded border" style={{ borderColor: 'var(--rule)' }}>
                            <div className="flex justify-between items-center text-[10px] mb-1.5">
                              <span className="text-slate-500">Active usage/month:</span>
                              <span className="font-bold font-mono">{sub.usage} days</span>
                            </div>
                            <input
                              type="range"
                              min="0"
                              max="30"
                              className="w-full h-1 rounded-lg appearance-none cursor-pointer accent-[var(--ink)]"
                              style={{ background: 'var(--paper-deep)' }}
                              value={sub.usage}
                              onChange={(e) => {
                                const nextUsage = { ...subscriptionUsage, [sub.name]: parseInt(e.target.value) };
                                setSubscriptionUsage(nextUsage);
                              }}
                            />
                            {sub.isLowUsage && (
                              <div className="mt-2 text-[9px] font-bold text-[var(--stamp)]">
                                <Icons.AlertWarning className="w-4 h-4 inline mr-1 text-[var(--stamp)]" /> Low usage warning. Cost per usage day is excessive ({fmt(sub.amount / Math.max(1, sub.usage))}).
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* Slider projection card */}
              <div className="blur-card rounded p-5 space-y-4">
                <div>
                  <h3 className="font-bold text-xs uppercase tracking-wider text-slate-800">50% Discretionary Spend Redirector</h3>
                  <p className="text-[11px] text-slate-500 mt-0.5">Simulate saving by cooking at home or cutting Swiggy orders</p>
                </div>
                
                <div className="flex items-center gap-4 bg-slate-900/5 p-4 rounded border" style={{ borderColor: 'var(--rule)' }}>
                  <span className="text-xs font-bold w-12 text-slate-700">Cut: {reductionPercent}%</span>
                  <input
                    type="range"
                    min="10"
                    max="90"
                    step="5"
                    className="flex-1 h-1.5 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-[var(--ink)]"
                    value={reductionPercent}
                    onChange={e => setReductionPercent(parseInt(e.target.value))}
                  />
                  <span className="text-xs font-bold text-emerald-800 font-mono">Redirect: {fmt(discretionarySaved)}/mo</span>
                </div>

                <div className="grid grid-cols-2 gap-4 text-center">
                  <div className="border rounded p-3" style={{ borderColor: 'var(--rule)', background: 'rgba(255,255,255,0.3)' }}>
                    <span className="text-[10px] text-slate-500 uppercase block font-bold">5-Year Wealth (12% compounding)</span>
                    <span className="text-lg font-bold font-mono text-[var(--positive)]">{fmt(wealthLossProjection5Yr)}</span>
                  </div>
                  <div className="border rounded p-3" style={{ borderColor: 'var(--rule)', background: 'rgba(255,255,255,0.3)' }}>
                    <span className="text-[10px] text-slate-500 uppercase block font-bold">10-Year Wealth (12% compounding)</span>
                    <span className="text-lg font-bold font-mono text-[var(--positive)]">{fmt(wealthLossProjection10Yr)}</span>
                  </div>
                </div>
              </div>

            </div>
          )}

          {/* ============================================================== */}
          {/* TAB 4: GOALS & ROOM DUES MANAGER */}
          {/* ============================================================== */}
          {activeTab === 'goals' && (
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              
              {/* Roommate Management & Dues */}
              <div className="blur-card rounded p-5 lg:col-span-2 space-y-6">
                <div>
                  <h3 className="font-bold text-xs uppercase tracking-wider text-slate-950 flex justify-between items-center border-b pb-2" style={{ borderColor: 'var(--rule)' }}>
                    <span>👥 Roommate flatmates list</span>
                    <span className="text-[10px] font-normal lowercase text-slate-500">invite by email to link them</span>
                  </h3>
                  
                  {/* List roommates */}
                  <div className="space-y-2 mt-3">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 p-2.5 bg-slate-900/5 rounded border text-xs" style={{ borderColor: 'var(--rule)' }}>
                      <div className="min-w-0">
                        <span className="font-bold text-slate-800 truncate block">{currentUser?.name} (You)</span>
                        <span className="text-[9px] text-slate-500 block truncate">{currentUser?.email}</span>
                      </div>
                      <div className="flex-shrink-0">
                        <span className={`px-1.5 py-0.5 font-bold border text-[9px] uppercase rounded ${
                          roomAdminId === session?.user?.id 
                            ? 'bg-amber-600/10 text-amber-800 border-amber-500/20' 
                            : 'bg-slate-900/5 text-slate-600 border-slate-500/10'
                        }`}>
                          {roomAdminId === session?.user?.id ? 'Group Admin' : 'Member'}
                        </span>
                      </div>
                    </div>

                    {roommates.map(member => {
                      const isAdmin = member.id === roomAdminId;
                      return (
                        <div key={member.id} className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 p-2.5 bg-slate-900/5 rounded border text-xs" style={{ borderColor: 'var(--rule)' }}>
                          <div className="min-w-0">
                            <span className="font-bold text-slate-800 truncate block">{member.name}</span>
                            <span className="text-[9px] text-slate-500 block truncate">{member.email}</span>
                          </div>
                          <div className="flex items-center space-x-2 flex-shrink-0">
                            <span className={`px-1.5 py-0.5 font-bold border text-[9px] uppercase rounded ${
                              isAdmin 
                                ? 'bg-amber-600/10 text-amber-800 border-amber-500/20' 
                                : 'bg-slate-900/5 text-slate-600 border-slate-500/10'
                            }`}>
                              {isAdmin ? 'Group Admin' : 'Member'}
                            </span>
                            {roomAdminId === session?.user?.id && !isAdmin && (
                              <button 
                                onClick={() => {
                                  if (confirm(`Are you sure you want to transfer Group Admin role to ${member.name}?`)) {
                                    transferAdminRole(member.id, member.name);
                                  }
                                }}
                                className="px-2 py-0.5 bg-[var(--ink)] text-[var(--card)] rounded text-[9px] font-bold"
                              >
                                Make Admin
                              </button>
                            )}
                          </div>
                        </div>
                      );
                    })}
                    
                    {roommates.length === 0 && (
                      <div className="text-center py-4 text-xs text-slate-500">
                        No roommates linked yet {currentUser?.name?.split(' ')[0] || ''}. Click "Add Flatmate" to send an invitation by email.
                      </div>
                    )}
                  </div>

                  {/* Pending join requests list (only shown to Admin) */}
                  {roomAdminId === session?.user?.id && pendingMembers.length > 0 && (
                    <div className="mt-4 pt-4 border-t" style={{ borderColor: 'var(--rule)' }}>
                      <h4 className="text-[10px] uppercase font-bold tracking-wider text-[var(--stamp)] mb-2 flex items-center gap-1.5">
                        <span>📩 Incoming flat join requests:</span>
                      </h4>
                      <div className="space-y-2">
                        {pendingMembers.map(member => (
                          <div key={member.id} className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 p-2.5 bg-amber-600/5 rounded border border-amber-500/20 text-xs animate-fade-in">
                            <div className="min-w-0">
                              <span className="font-bold text-slate-800 truncate block">{member.name}</span>
                              <span className="text-[9px] text-slate-500 block truncate">{member.email}</span>
                            </div>
                            <div className="flex items-center gap-1.5 flex-shrink-0">
                              <button 
                                onClick={() => handleApproveRoommateRequest(member.id)}
                                className="px-2.5 py-1 bg-emerald-700 hover:bg-emerald-800 text-white rounded text-[10px] font-bold transition shadow-xs"
                              >
                                Approve
                              </button>
                              <button 
                                onClick={async () => {
                                  if (confirm(`Decline and remove request from ${member.name}?`)) {
                                    try {
                                      await supabase
                                        .from('room_members')
                                        .delete()
                                        .match({ room_id: currentRoomId, user_id: member.id });
                                      showToast('success', 'Join request declined.');
                                      await fetchRoommatesAndTransactions(session.user.id, currentRoomId);
                                    } catch (e) {
                                      showToast('error', 'Error declining request.');
                                    }
                                  }
                                }}
                                className="px-2.5 py-1 border border-slate-300 hover:bg-slate-100 text-slate-600 rounded text-[10px] font-bold transition"
                              >
                                Decline
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* Dues Settlement Matrix */}
                <div>
                  <h3 className="font-bold text-xs uppercase tracking-wider text-slate-950 mb-3 border-b pb-2" style={{ borderColor: 'var(--rule)' }}>
                    💵 Outstanding splits settlements
                  </h3>
                  
                  {roommateDues.duesList.length === 0 ? (
                    <div className="text-xs text-slate-500 text-center py-4">No outstanding balances between roommate group! Everyone is squared up.</div>
                  ) : (
                    <div className="space-y-2.5">
                      {roommateDues.duesList.map((due, idx) => (
                        <div key={idx} className="flex flex-col sm:flex-row sm:items-center justify-between gap-3.5 p-3 bg-white rounded border hover:shadow-sm transition" style={{ borderColor: 'var(--rule)' }}>
                          <div className="text-xs flex flex-wrap items-center gap-1.5 min-w-0">
                            <span className="font-bold text-red-700 truncate max-w-[120px]">{due.from}</span>
                            <span className="text-slate-400">owes</span>
                            <span className="font-bold text-emerald-800 truncate max-w-[120px]">{due.to}</span>
                            <span className="font-bold font-mono bg-slate-100 px-2 py-0.5 rounded flex-shrink-0">{fmt(due.amount)}</span>
                          </div>
                          
                          {/* Settle button */}
                          <div className="flex-shrink-0 w-full sm:w-auto text-right">
                            {roomAdminId === session?.user?.id ? (
                              <button 
                                onClick={() => handleClearDuesDirect(due.from, due.to, due.amount)}
                                className="w-full sm:w-auto px-3 py-1 bg-[var(--ink)] text-[var(--card)] rounded text-[10px] font-bold"
                              >
                                Mark Paid
                              </button>
                            ) : (
                              <span className="text-[10px] text-slate-400 italic font-medium block">Admin only</span>
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              {/* Personal Goals (Right column) */}
              <div className="blur-card rounded p-5 lg:col-span-1 space-y-4">
                <h3 className="font-bold text-xs uppercase tracking-wider text-slate-950 border-b pb-2" style={{ borderColor: 'var(--rule)' }}>🎯 {currentUser?.name?.split(' ')[0] || 'My'}'s Bachelor Goals</h3>
                <div className="space-y-4">
                  {goals.map(g => {
                    const percent = Math.min(100, Math.round((g.current / g.target) * 100));
                    return (
                      <div key={g.id} className="space-y-1">
                        <div className="flex justify-between text-xs font-bold">
                          <span className="text-slate-800">{g.name}</span>
                          <span className="font-mono">{percent}%</span>
                        </div>
                        <div className="w-full h-2 bg-slate-200 rounded-full overflow-hidden border border-slate-300">
                          <div 
                            className="h-full bg-[var(--positive)] transition-all duration-300"
                            style={{ width: `${percent}%` }}
                          />
                        </div>
                        <div className="flex justify-between text-[10px] text-slate-500 font-mono">
                          <span>{fmt(g.current)} saved</span>
                          <span>Target: {fmt(g.target)}</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

            </div>
          )}

          {/* ============================================================== */}
          {/* TAB 5: AI COACH & FORECASTING */}
          {/* ============================================================== */}
          {activeTab === 'coach' && (
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              
              {/* Chat Window */}
              <div className="blur-card rounded p-5 lg:col-span-2 flex flex-col justify-between h-[450px]">
                <div>
                  <h3 className="font-bold text-xs uppercase tracking-wider text-slate-900 border-b pb-2 mb-3" style={{ borderColor: 'var(--rule)' }}>
                    💬 AI Bachelor Financial Coach
                  </h3>
                  
                  {/* Messages Feed */}
                  <div className="h-[310px] overflow-y-auto space-y-3 pr-1 text-xs">
                    {chatMessages.map(msg => (
                      <div key={msg.id} className={`flex ${msg.sender === 'user' ? 'justify-end' : 'justify-start'}`}>
                        <div 
                          className={`max-w-[80%] p-3 rounded shadow-sm whitespace-pre-line leading-relaxed ${msg.sender === 'user' ? 'bg-[var(--ink)] text-[var(--card)] rounded-br-none' : 'bg-slate-900/5 text-slate-800 rounded-bl-none border'}`}
                          style={{ borderColor: msg.sender !== 'user' ? 'var(--rule)' : 'transparent' }}
                        >
                          <p className="font-bold text-[9px] uppercase tracking-wider mb-1 opacity-70">
                            {msg.sender === 'user' ? currentUser?.name : 'AI coach'}
                          </p>
                          {msg.text}
                        </div>
                      </div>
                    ))}
                    <div ref={chatEndRef} />
                  </div>
                </div>

                {/* Form Chat input */}
                <div className="flex gap-2 border-t pt-3 mt-3" style={{ borderColor: 'var(--rule)' }}>
                  <input
                    type="text"
                    placeholder="Ask about leaks, roommate split check, or survival pace..."
                    className="flex-1 ledger-input-box"
                    value={chatInput}
                    onChange={e => setChatInput(e.target.value)}
                    onKeyDown={e => { if (e.key === 'Enter') askAI(chatInput); }}
                  />
                  <button onClick={() => askAI(chatInput)} className="btn-vintage-ink flex-shrink-0">Send</button>
                </div>
              </div>

              {/* Dynamic Mistake Flags */}
              <div className="blur-card rounded p-5 lg:col-span-1 flex flex-col justify-between">
                <div>
                  <h3 className="font-bold text-xs uppercase tracking-wider text-slate-900 border-b pb-2 mb-3" style={{ borderColor: 'var(--rule)' }}>
                    🚨 Heuristic Mistake flags
                  </h3>
                  
                  {mistakeAlerts.length === 0 ? (
                    <div className="text-xs text-slate-500 py-4 text-center">No budget leakages flagged. Keep up the high savings rate!</div>
                  ) : (
                    <div className="space-y-3">
                      {mistakeAlerts.map((alert, idx) => (
                        <div key={idx} className="p-3 bg-red-600/5 border border-red-500/20 rounded flex items-start gap-2">
                          <Icons.Siren className="w-4 h-4 text-red-600 flex-shrink-0 mt-0.5" />
                          <div>
                            <h4 className="font-bold text-xs text-red-900">{alert.title}</h4>
                            <p className="text-[10px] text-red-800/80 mt-0.5 leading-relaxed">{alert.desc}</p>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                <div className="text-[10px] text-slate-500 pt-3 border-t mt-4" style={{ borderColor: 'var(--rule)' }}>
                  Flags compile automatically based on recurring discretionary spend patterns.
                </div>
              </div>

            </div>
          )}

          {/* ============================================================== */}
          {/* TAB 6: STATEMENT & INGESTION HUB */}
          {/* ============================================================== */}
          {activeTab === 'ingestion' && (
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              
              {/* CSV Upload */}
              <div className="blur-card rounded p-5 lg:col-span-2 flex flex-col justify-between min-h-[300px]">
                <div>
                  <h3 className="font-bold text-xs uppercase tracking-wider text-slate-900 border-b pb-2 mb-4" style={{ borderColor: 'var(--rule)' }}>
                    📂 Bank Statement CSV Parser
                  </h3>
                  
                  <div className="border-2 border-dashed rounded-lg p-6 text-center space-y-3 cursor-pointer hover:bg-slate-900/5 transition" style={{ borderColor: 'var(--rule)' }}>
                    <Icons.Ingestion className="w-8 h-8 mx-auto text-slate-400" />
                    <div className="text-xs text-slate-600">
                      Drag & drop your bank statement CSV file or <span className="underline font-bold text-[var(--ink)]">browse files</span>
                    </div>
                    <input 
                      type="file" 
                      accept=".csv"
                      className="hidden" 
                      id="statementFileUploader"
                      onChange={handleCsvUpload}
                    />
                    <label htmlFor="statementFileUploader" className="px-3 py-1.5 bg-[var(--ink)] text-[var(--card)] rounded text-[10px] font-bold cursor-pointer inline-block">
                      Choose File
                    </label>
                  </div>

                  {/* CSV Preview */}
                  {csvFilePreview && (
                    <div className="mt-4 space-y-3">
                      <div className="flex justify-between items-center text-xs bg-slate-900/5 p-2 rounded">
                        <span className="font-bold text-slate-800">Previewing: {csvFilePreview.name} ({csvFilePreview.size})</span>
                        <button onClick={importCsvPreviewRows} className="px-2.5 py-1 bg-emerald-600 text-white rounded text-[10px] font-bold hover:bg-emerald-700">Import Rows</button>
                      </div>

                      <div className="border rounded overflow-hidden" style={{ borderColor: 'var(--rule)' }}>
                        <table className="w-full text-[10px] text-left">
                          <thead className="bg-slate-100 text-slate-600 font-bold border-b" style={{ borderColor: 'var(--rule)' }}>
                            <tr>
                              <th className="p-2">Date</th>
                              <th className="p-2">Merchant</th>
                              <th className="p-2">Category</th>
                              <th className="p-2">Amount</th>
                              <th className="p-2">Split flat?</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y font-mono text-slate-700" style={{ borderColor: 'var(--rule)' }}>
                            {csvFilePreview.rows.map((row, idx) => (
                              <tr key={idx}>
                                <td className="p-2">{row.date}</td>
                                <td className="p-2 font-sans font-bold">{row.merchant}</td>
                                <td className="p-2 font-sans">{row.category}</td>
                                <td className="p-2 font-bold">{fmt(row.amount)}</td>
                                <td className="p-2 text-emerald-800 font-bold">{row.split}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>

                <div className="text-[10px] text-slate-400 pt-3 border-t mt-4" style={{ borderColor: 'var(--rule)' }}>
                  All bank files analyze securely in-browser. No data uploads to external cloud processors.
                </div>
              </div>

              {/* Receipt OCR and Voice entry Column */}
              <div className="space-y-6 lg:col-span-1">
                
                {/* Receipt OCR */}
                <div className="blur-card rounded p-5 space-y-3">
                  <h3 className="font-bold text-xs uppercase tracking-wider text-slate-900 border-b pb-1.5" style={{ borderColor: 'var(--rule)' }}>
                    📷 Receipt OCR Scanner
                  </h3>
                  
                  <div className="border border-dashed rounded p-4 text-center cursor-pointer hover:bg-slate-900/5 transition relative overflow-hidden" style={{ borderColor: 'var(--rule)' }}>
                    <Icons.Camera className="w-6 h-6 mx-auto text-slate-400 mb-1" />
                    <span className="text-[10px] text-slate-500 block">Scan receipt image (PNG/JPG)</span>
                    <input 
                      type="file" 
                      accept="image/*"
                      className="hidden" 
                      id="receiptFileUploader"
                      onChange={handleReceiptOcrScan}
                    />
                    <label htmlFor="receiptFileUploader" className="mt-2 px-2 py-1 border border-[var(--ink)] text-[var(--ink)] rounded text-[9px] font-bold cursor-pointer inline-block">
                      Upload Receipt
                    </label>

                    {ocrScanning && (
                      <div className="absolute inset-0 bg-slate-950/20 flex flex-col items-center justify-center">
                        <div className="w-full h-0.5 bg-emerald-500 absolute scanner-beam" />
                        <span className="text-[10px] text-slate-950 bg-white px-2 py-0.5 rounded font-bold shadow-sm">Scanning...</span>
                      </div>
                    )}
                  </div>

                  {ocrResult && (
                    <div className="bg-emerald-600/5 border border-emerald-500/20 p-3 rounded text-[11px] space-y-1.5">
                      <div className="flex justify-between font-bold text-slate-800">
                        <span>Merchant: {ocrResult.merchant}</span>
                        <span className="text-emerald-800">Confidence: {ocrResult.confidence}</span>
                      </div>
                      <div className="flex justify-between font-mono text-slate-600">
                        <span>Date: {ocrResult.date}</span>
                        <span className="font-sans font-bold text-slate-800">Spent: {fmt(ocrResult.amount)}</span>
                      </div>
                      <button onClick={acceptOcrTransaction} className="w-full py-1 bg-emerald-600 text-white rounded text-[9px] font-bold hover:bg-emerald-700">Approve & log expense</button>
                    </div>
                  )}
                </div>

                {/* Voice Entry */}
                <div className="blur-card rounded p-5 space-y-3">
                  <h3 className="font-bold text-xs uppercase tracking-wider text-slate-900 border-b pb-1.5" style={{ borderColor: 'var(--rule)' }}>
                    🎙️ Voice Command Entry
                  </h3>
                  
                  <div className="flex gap-2">
                    <button 
                      onClick={triggerVoiceRecording}
                      className={`w-10 h-10 rounded-full flex items-center justify-center flex-shrink-0 transition-all border ${voiceRecording ? 'bg-red-600 text-white animate-pulse border-red-700' : 'bg-slate-900/5 text-slate-600 hover:bg-slate-900/10'}`}
                      style={{ borderColor: !voiceRecording ? 'var(--rule)' : '' }}
                    >
                      <Icons.Microphone className="w-4 h-4" />
                    </button>
                    
                    <div className="min-w-0 flex-1 flex flex-col justify-center">
                      <span className="text-[10px] font-bold text-slate-500 uppercase block">Record intent Command</span>
                      <span className="text-[9px] text-slate-400 truncate">
                        {voiceRecording ? `Listening ${currentUser?.name?.split(' ')[0] || ''}...` : voiceText || 'Click microphone to dictate spent'}
                      </span>
                    </div>
                  </div>

                  {voiceResult && (
                    <div className="bg-slate-900/5 border p-3 rounded text-[11px] space-y-1.5" style={{ borderColor: 'var(--rule)' }}>
                      <p className="font-semibold text-slate-700">Heuristics Parse: "{voiceResult.summary}"</p>
                      <button onClick={acceptVoiceTransaction} className="w-full py-1 bg-[var(--ink)] text-[var(--card)] rounded text-[9px] font-bold">Approve & Log entry</button>
                    </div>
                  )}
                </div>

              </div>

            </div>
          )}
        </div>

        {/* Roommates Mode Overlays (when not set up or pending) */}
        {analysisType === 'roommates' && roomMembershipStatus === 'none' && (
          <div className="flex-1 flex items-center justify-center p-4 my-auto">
            <style>{`
              .room-card {
                background: var(--card);
                border: 1px solid var(--rule);
                box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
              }
              .btn-vintage-outline {
                background: transparent;
                border: 1px solid var(--ink);
                color: var(--ink);
                font-weight: 600;
                font-size: 13px;
                padding: 9px 16px;
                border-radius: 4px;
                width: 100%;
                transition: background 0.15s;
              }
              .btn-vintage-outline:hover {
                background: rgba(30,42,68,0.05);
              }
            `}</style>
            <div className="room-card max-w-md w-full rounded-lg p-6 space-y-6">
              <div className="text-center space-y-1">
                <h2 className="text-lg font-bold text-[var(--ink)]">Setup Your Flat Roommates Group</h2>
                <p className="text-[10px] uppercase tracking-wider text-[var(--ink-soft)] font-bold">Connect with roommates to share splits</p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="border p-4 rounded flex flex-col justify-between space-y-3 animate-fade-in" style={{ borderColor: 'var(--rule)' }}>
                  <div>
                    <h3 className="font-bold text-xs text-slate-800 uppercase tracking-wider">Option A: Create Room</h3>
                    <p className="text-[11px] text-slate-500 mt-1">Start a new roommates flat group. You will get an invite code to share with Sriram, Kartik, or Arul.</p>
                  </div>
                  <button onClick={handleCreateRoom} className="btn-vintage-ink mt-3" disabled={roomLoading}>
                    {roomLoading ? "Creating..." : "Create New Flat"}
                  </button>
                </div>

                <div className="border p-4 rounded flex flex-col justify-between space-y-3 animate-fade-in" style={{ borderColor: 'var(--rule)' }}>
                  <div>
                    <h3 className="font-bold text-xs text-slate-800 uppercase tracking-wider">Option B: Join Room</h3>
                    <p className="text-[11px] text-slate-500 mt-1">Enter your flatmates' shared Invite Code to join their existing flat roommate sheets.</p>
                  </div>
                  <div className="space-y-2">
                    <input
                      type="text"
                      placeholder="Invite Code (e.g. FLAT-xxxx)"
                      className="ledger-input-box uppercase"
                      value={inviteCodeInput}
                      onChange={e => setInviteCodeInput(e.target.value.toUpperCase())}
                    />
                    <button onClick={handleJoinRoom} className="btn-vintage-outline" disabled={roomLoading}>
                      {roomLoading ? "Submitting..." : "Send Join Request"}
                    </button>
                  </div>
                </div>
              </div>

              {pendingInvites.length > 0 && (
                <div className="border-t pt-4 space-y-2" style={{ borderColor: 'var(--rule)' }}>
                  <h4 className="text-[10px] uppercase font-bold tracking-wider text-[var(--ink-soft)]">📩 Pending invitations from roommates:</h4>
                  <div className="space-y-1.5">
                    {pendingInvites.map(invite => (
                      <div key={invite.room_id} className="flex justify-between items-center p-2 bg-slate-900/5 rounded border" style={{ borderColor: 'var(--rule)' }}>
                        <div>
                          <p className="text-xs font-bold text-slate-800">Flat Invite (Code: {invite.invite_code})</p>
                          <p className="text-[10px] text-slate-500">Sent on: {new Date(invite.joined_at).toLocaleDateString()}</p>
                        </div>
                        <div className="flex gap-1.5">
                          <button onClick={() => handleAcceptInvite(invite.room_id)} className="px-2 py-0.5 bg-[var(--ink)] text-[var(--card)] rounded text-[10px] font-bold">Accept</button>
                          <button onClick={() => handleRejectInvite(invite.room_id)} className="px-2 py-0.5 border border-[var(--stamp)] text-[var(--stamp)] rounded text-[10px] font-bold">Ignore</button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {errorMsg && (
                <p className="text-xs text-[var(--stamp)] font-bold text-center">{errorMsg}</p>
              )}
            </div>
          </div>
        )}

        {analysisType === 'roommates' && currentRoomId && roomMembershipStatus === 'pending' && (
          <div className="flex-1 flex items-center justify-center p-4 my-auto">
            <div className="room-card max-w-sm w-full rounded-lg p-6 space-y-6 text-center">
              <Icons.AddUser className="w-10 h-10 mx-auto text-[var(--ink-soft)]" />
              <h2 className="text-lg font-bold text-[var(--ink)]">Membership Pending</h2>
              <p className="text-xs text-slate-600">
                Your request to join Room Invite Code: <strong>{currentRoomCode}</strong> has been sent. 
              </p>
              <p className="text-xs text-slate-500 bg-slate-900/5 p-2 rounded mt-2">
                Ask the roommates who are already inside the flat to click <strong>"Accept"</strong> on their dashboard.
              </p>
              <div className="flex flex-col gap-2 pt-2">
                <button onClick={checkMembershipStatus} className="btn-vintage-ink">Check Status Again</button>
                <button onClick={handleLeaveRoom} className="btn-vintage-outline">Cancel Request / Join Another Flat</button>
              </div>
            </div>
          </div>
        )}

        {/* Footer Privacy statement */}
        <footer className="mt-8 pt-4 border-t flex flex-col sm:flex-row items-center justify-between text-[9px] text-slate-400 gap-2" style={{ borderColor: 'var(--rule)' }}>
          <span className="flex items-center gap-1">
            <Icons.Shield className="w-3.5 h-3.5 text-emerald-600" /> Local database privacy shield active (Vite + Supabase connection secure)
          </span>
          <span>© 2026 FinSense AI Collaborative Bachelor Ledger</span>
        </footer>

      </main>

      </div> {/* Closing Main Inner Flex Wrapper */}

      {/* Roommate email invitation popup modal */}
      {isAuthOpen && (
        <div className="fixed inset-0 bg-slate-950/40 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="room-card max-w-sm w-full rounded-lg p-5 space-y-4" style={{ background: 'var(--card)', border: '1px solid var(--rule)' }}>
            <div className="flex justify-between items-center border-b pb-2" style={{ borderColor: 'var(--rule)' }}>
              <h3 className="font-bold text-xs uppercase tracking-wider text-[var(--ink)]">Add Flat Roommate</h3>
              <button onClick={() => setIsAuthOpen(false)} className="text-xs font-bold text-slate-400 hover:text-slate-800">✕</button>
            </div>
            
            <form onSubmit={handleAddRoommate} className="space-y-3.5">
              <div className="flex flex-col">
                <label className="text-[9px] uppercase font-bold tracking-wider mb-1 text-slate-500">Roommate Email Address</label>
                <input
                  type="email"
                  placeholder="e.g. karthik@room.com"
                  className="ledger-input-box"
                  value={authForm.email}
                  onChange={e => setAuthForm({ ...authForm, email: e.target.value })}
                  required
                />
              </div>
              <p className="text-[10px] text-slate-500 leading-relaxed">
                Enter your roommate's registered email. They will receive a pending request on their home dashboard and join your shared split group once they click Accept.
              </p>
              
              {roomLoading ? (
                <div className="text-xs text-center text-slate-500">Processing invitation...</div>
              ) : (
                <div className="flex gap-2">
                  <button type="submit" className="flex-1 btn-vintage-ink">Send Invitation</button>
                  <button type="button" onClick={() => setIsAuthOpen(false)} className="flex-1 btn-vintage-outline">Cancel</button>
                </div>
              )}
            </form>
          </div>
        </div>
      )}

    </div>
  );
}


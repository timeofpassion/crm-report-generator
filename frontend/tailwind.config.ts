import type { Config } from 'tailwindcss';

const config: Config = {
  content: ['./src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        navy: '#1A3A6B',
        blue: { DEFAULT: '#2D7DD2', light: '#EBF3FB' },
        amber: '#E8A020',
      },
      fontFamily: {
        sans: ['Pretendard', 'Apple SD Gothic Neo', 'Malgun Gothic', 'sans-serif'],
      },
    },
  },
};

export default config;

/**
 * Injects custom CSS styles directly into the document head
 * This avoids webpack CSS processing issues
 */
export function injectStyles(): () => void {
  if (typeof document === 'undefined') {
    // Return a no-op cleanup function when document is not available
    return () => {};
  }

  const styleElement = document.createElement('style');
  styleElement.textContent = `
    /* Base styles for the Excel add-in - matching cascade-chat */
    :root {
      --background-color: #050508;
      --foreground-color: #ffffff;
      --primary-color: rgba(76, 194, 255, 1);
      --secondary-color: rgba(40, 40, 40, 0.7);
      --border-color: rgba(17, 24, 39, 0.3);
      --accent-color: rgba(76, 194, 255, 0.15);
    }

    body {
      color: var(--foreground-color);
      background: linear-gradient(to right bottom, rgba(0, 0, 0, 0.95), rgba(20, 20, 20, 0.9)), 
                  radial-gradient(circle at 50% 0%, rgba(76, 194, 255, 0.3), transparent 70%);
      background-attachment: fixed;
      background-size: cover;
      margin: 0;
      padding: 0;
      min-height: 100vh;
      font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
      font-size: 14px;
      line-height: 1.5;
      overflow: hidden;
    }

    /* Typography classes similar to cascade-chat */
    .text-xs { font-size: 0.75rem; line-height: 1rem; }
    .text-sm { font-size: 0.875rem; line-height: 1.25rem; }
    .text-base { font-size: 1rem; line-height: 1.5rem; }
    .text-lg { font-size: 1.125rem; line-height: 1.75rem; }
    .text-xl { font-size: 1.25rem; line-height: 1.75rem; }
    .font-normal { font-weight: 400; }
    .font-medium { font-weight: 500; }
    .font-semibold { font-weight: 600; }
    .font-bold { font-weight: 700; }
    .leading-normal { line-height: 1.5; }
    .leading-relaxed { line-height: 1.625; }
    .tracking-tight { letter-spacing: -0.025em; }
    
    /* Color classes */
    .text-white { color: white; }
    .text-white\/30 { color: rgba(255, 255, 255, 0.3); }
    .text-white\/60 { color: rgba(255, 255, 255, 0.6); }
    .text-white\/70 { color: rgba(255, 255, 255, 0.7); }
    .text-white\/80 { color: rgba(255, 255, 255, 0.8); }
    .text-gray-300 { color: rgb(209, 213, 219); }
    .text-gray-400 { color: rgb(156, 163, 175); }
    .text-green-400 { color: rgb(74, 222, 128); }
    .bg-transparent { background-color: transparent; }
    .bg-black\/50 { background-color: rgba(0, 0, 0, 0.5); }
    .bg-black\/70 { background-color: rgba(0, 0, 0, 0.7); }
    .bg-blue-600 { background-color: rgb(37, 99, 235); }
    .border-gray-900\/30 { border-color: rgba(17, 24, 39, 0.3); }
    .border-accent { border-color: var(--accent-color); }
    
    /* Layout classes */
    .flex { display: flex; }
    .inline-flex { display: inline-flex; }
    .flex-col { flex-direction: column; }
    .flex-1 { flex: 1 1 0%; }
    .flex-grow { flex-grow: 1; }
    .flex-shrink-0 { flex-shrink: 0; }
    .items-start { align-items: flex-start; }
    .items-center { align-items: center; }
    .items-end { align-items: flex-end; }
    .justify-start { justify-content: flex-start; }
    .justify-center { justify-content: center; }
    .justify-between { justify-content: space-between; }
    .justify-end { justify-content: flex-end; }
    .gap-1 { gap: 0.25rem; }
    .gap-2 { gap: 0.5rem; }
    .gap-3 { gap: 0.75rem; }
    .gap-4 { gap: 1rem; }

    /* Spacing - matching cascade-chat */
    .p-1 { padding: 0.25rem; }
    .p-2 { padding: 0.5rem; }
    .p-3 { padding: 0.75rem; }
    .p-4 { padding: 1rem; }
    .p-5 { padding: 1.25rem; }
    .p-7 { padding: 1.75rem; }
    .px-1 { padding-left: 0.25rem; padding-right: 0.25rem; }
    .px-2 { padding-left: 0.5rem; padding-right: 0.5rem; }
    .px-3 { padding-left: 0.75rem; padding-right: 0.75rem; }
    .px-4 { padding-left: 1rem; padding-right: 1rem; }
    .py-0\.5 { padding-top: 0.125rem; padding-bottom: 0.125rem; }
    .py-1 { padding-top: 0.25rem; padding-bottom: 0.25rem; }
    .py-2 { padding-top: 0.5rem; padding-bottom: 0.5rem; }
    .py-3 { padding-top: 0.75rem; padding-bottom: 0.75rem; }
    .py-4 { padding-top: 1rem; padding-bottom: 1rem; }
    .pt-1 { padding-top: 0.25rem; }
    .pt-2 { padding-top: 0.5rem; }
    .pb-2 { padding-bottom: 0.5rem; }
    .pl-2 { padding-left: 0.5rem; }
    .pr-2 { padding-right: 0.5rem; }
    
    .m-0 { margin: 0; }
    .m-1 { margin: 0.25rem; }
    .m-2 { margin: 0.5rem; }
    .m-3 { margin: 0.75rem; }
    .m-4 { margin: 1rem; }
    .mt-0 { margin-top: 0; }
    .mt-1 { margin-top: 0.25rem; }
    .mt-2 { margin-top: 0.5rem; }
    .mt-3 { margin-top: 0.75rem; }
    .mt-4 { margin-top: 1rem; }
    .mt-8 { margin-top: 2rem; }
    .mb-0 { margin-bottom: 0; }
    .mb-1 { margin-bottom: 0.25rem; }
    .mb-2 { margin-bottom: 0.5rem; }
    .mb-3 { margin-bottom: 0.75rem; }
    .mb-4 { margin-bottom: 1rem; }
    .mb-7 { margin-bottom: 1.75rem; }
    .mb-10 { margin-bottom: 2.5rem; }
    .mb-14 { margin-bottom: 3.5rem; }
    .ml-1 { margin-left: 0.25rem; }
    .ml-2 { margin-left: 0.5rem; }
    .ml-auto { margin-left: auto; }
    .mr-1 { margin-right: 0.25rem; }
    .mr-2 { margin-right: 0.5rem; }
    .mr-auto { margin-right: auto; }
    
    /* Sizing */
    .h-2 { height: 0.5rem; }
    .h-3 { height: 0.75rem; }
    .h-4 { height: 1rem; }
    .h-6 { height: 1.5rem; }
    .h-8 { height: 2rem; }
    .h-10 { height: 2.5rem; }
    .h-14 { height: 3.5rem; }
    .h-screen { height: 100vh; }
    .h-full { height: 100%; }
    .w-2 { width: 0.5rem; }
    .w-3 { width: 0.75rem; }
    .w-4 { width: 1rem; }
    .w-6 { width: 1.5rem; }
    .w-8 { width: 2rem; }
    .w-10 { width: 2.5rem; }
    .w-14 { width: 3.5rem; }
    .w-full { width: 100%; }
    .w-max { width: max-content; }
    .w-screen { width: 100vw; }
    .max-w-xs { max-width: 20rem; }
    .max-w-sm { max-width: 24rem; }
    .max-w-md { max-width: 28rem; }
    .max-w-lg { max-width: 32rem; }
    .max-h-60 { max-height: 15rem; }
    .min-h-screen { min-height: 100vh; }
    
    /* Borders and Rounding */
    .rounded-sm { border-radius: 0.125rem; }
    .rounded { border-radius: 0.25rem; }
    .rounded-md { border-radius: 0.375rem; }
    .rounded-lg { border-radius: 0.5rem; }
    .rounded-full { border-radius: 9999px; }
    .border { border-width: 1px; }
    .border-0 { border-width: 0px; }
    .border-t { border-top-width: 1px; }
    .border-b { border-bottom-width: 1px; }
    .border-l { border-left-width: 1px; }
    .border-r { border-right-width: 1px; }
    .border-4 { border-width: 4px; }
    
    /* Effects */
    .glass-dark {
      background-color: rgba(40, 40, 40, 0.7);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      border: 1px solid rgba(76, 194, 255, 0.15);
    }

    .glass-darker {
      background-color: rgba(20, 20, 20, 0.8);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      border: 1px solid rgba(76, 194, 255, 0.15);
    }
    
    .shadow-sm { box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05); }
    .shadow { box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06); }
    .shadow-md { box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); }
    .shadow-lg { box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05); }
    
    /* Positioning */
    .relative { position: relative; }
    .absolute { position: absolute; }
    .fixed { position: fixed; }
    .sticky { position: sticky; }
    .top-0 { top: 0; }
    .top-1 { top: 0.25rem; }
    .top-2 { top: 0.5rem; }
    .top-1\/2 { top: 50%; }
    .right-0 { right: 0; }
    .right-1 { right: 0.25rem; }
    .right-2 { right: 0.5rem; }
    .right-3 { right: 0.75rem; }
    .bottom-0 { bottom: 0; }
    .bottom-1 { bottom: 0.25rem; }
    .bottom-2 { bottom: 0.5rem; }
    .left-0 { left: 0; }
    .left-1 { left: 0.25rem; }
    .left-2 { left: 0.5rem; }
    .-translate-y-1\/2 { transform: translateY(-50%); }
    .z-0 { z-index: 0; }
    .z-10 { z-index: 10; }
    .z-20 { z-index: 20; }
    .z-50 { z-index: 50; }

    /* Display & Overflow */
    .block { display: block; }
    .inline-block { display: inline-block; }
    .hidden { display: none; }
    .overflow-auto { overflow: auto; }
    .overflow-hidden { overflow: hidden; }
    .overflow-y-auto { overflow-y: auto; }
    .overflow-x-hidden { overflow-x: hidden; }
    .whitespace-pre-wrap { white-space: pre-wrap; }
    .whitespace-nowrap { white-space: nowrap; }
    
    /* Transitions & Animations */
    @keyframes pulse {
      0%, 100% { opacity: 1; }
      50% { opacity: 0.5; }
    }
    .animate-pulse { animation: pulse 1.5s cubic-bezier(0.4, 0, 0.6, 1) infinite; }
    .transition-all { transition-property: all; }
    .transition-colors { transition-property: background-color, border-color, color, fill, stroke; }
    .transition-opacity { transition-property: opacity; }
    .transition-transform { transition-property: transform; }
    .duration-150 { transition-duration: 150ms; }
    .duration-300 { transition-duration: 300ms; }
    .duration-500 { transition-duration: 500ms; }
    .ease-in-out { transition-timing-function: cubic-bezier(0.4, 0, 0.2, 1); }
    
    /* Interactive */
    .cursor-pointer { cursor: pointer; }
    .cursor-not-allowed { cursor: not-allowed; }
    .select-none { user-select: none; }
    .focus\:outline-none:focus { outline: 2px solid transparent; outline-offset: 2px; }
    .focus\:ring-2:focus { --tw-ring-offset-shadow: var(--tw-ring-inset) 0 0 0 var(--tw-ring-offset-width) var(--tw-ring-offset-color); --tw-ring-shadow: var(--tw-ring-inset) 0 0 0 calc(2px + var(--tw-ring-offset-width)) var(--tw-ring-color); box-shadow: var(--tw-ring-offset-shadow), var(--tw-ring-shadow), var(--tw-shadow, 0 0 #0000); }
    .hover\:bg-opacity-80:hover { --tw-bg-opacity: 0.8; }
    .hover\:text-white:hover { color: white; }
    .hover\:bg-black\/20:hover { background-color: rgba(0, 0, 0, 0.2); }
    
    /* Input specific */
    textarea, input {
      resize: none;
      outline: none;
      background-color: transparent;
      color: var(--foreground-color);
      width: 100%;
    }
    
    textarea::placeholder, input::placeholder {
      color: rgba(255, 255, 255, 0.5);
    }
    
    button {
      cursor: pointer;
      background-color: transparent;
      border: none;
      color: var(--foreground-color);
      display: inline-flex;
      align-items: center;
      justify-content: center;
    }
    
    button:disabled {
      cursor: not-allowed;
      opacity: 0.5;
    }
    
    /* Message styles */
    .message-user {
      background-color: rgba(40, 40, 40, 0.7);
      backdrop-filter: blur(12px);
      border-radius: 0.5rem;
      margin-bottom: 0.5rem;
      padding: 1rem;
      border: 1px solid rgba(76, 194, 255, 0.15);
    }
    
    .message-assistant {
      padding: 0.5rem;
      margin-bottom: 1.75rem;
    }
    
    /* Custom components */
    .header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 1rem;
      backdrop-filter: blur(12px);
      background-color: rgba(40, 40, 40, 0.7);
      border-bottom: 1px solid rgba(17, 24, 39, 0.3);
      z-index: 10;
    }
    
    .header-title {
      display: flex;
      align-items: center;
      color: white;
      font-weight: 500;
    }
    
    .input-container {
      position: relative;
      margin-top: 0.5rem;
      border-radius: 0.5rem;
      background-color: rgba(40, 40, 40, 0.7);
      backdrop-filter: blur(12px);
      border: 1px solid rgba(76, 194, 255, 0.15);
    }
    
    .action-button {
      display: inline-flex;
      align-items: center;
      padding: 0.25rem 0.5rem;
      border-radius: 0.25rem;
      transition: background-color 0.2s;
    }
    
    .action-button:hover {
      background-color: rgba(0, 0, 0, 0.2);
    }
  `;
  
  document.head.appendChild(styleElement);
  
  return () => {
    document.head.removeChild(styleElement);
  };
}

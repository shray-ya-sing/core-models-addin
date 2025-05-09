import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

/**
 * A utility function that merges multiple class names and handles Tailwind conflicts
 * @param inputs Array of class names or conditional class values
 * @returns A merged class string with Tailwind conflicts resolved
 */
export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

/**
 * Utility function to merge multiple class names including conditional ones
 * This is for backward compatibility with existing components
 */
export function cx(...classes: (string | Record<string, boolean> | undefined)[]) {
  // Filter out undefined values
  const validClasses = classes.filter(Boolean) as (string | Record<string, boolean>)[];
  
  // Process each class or class object
  const processedClasses = validClasses.flatMap(cls => {
    if (typeof cls === 'string') {
      return cls; // Return string class as is
    } else if (typeof cls === 'object') {
      // For objects, only include keys where value is truthy
      return Object.keys(cls).filter(key => cls[key]);
    }
    return '';
  });
  
  // Use our new cn function to merge classes
  return cn(processedClasses);
}

/**
 * Converts Tailwind-style classes to Fluent UI compatible styles
 * @param className The Tailwind-inspired class string
 * @returns An object with CSS properties
 */
export function tailwindToFluentStyles(className: string): React.CSSProperties {
  const styles: React.CSSProperties = {};
  
  // Parse the className string and convert to React inline styles
  // This is a simplified implementation - extend as needed
  
  // Glass effect classes
  if (className.includes('glass-dark')) {
    styles.backgroundColor = 'rgba(26, 26, 26, 0.7)';
    styles.backdropFilter = 'blur(10px)';
    styles.borderRadius = '8px';
  }
  
  if (className.includes('glass-darker')) {
    styles.backgroundColor = 'rgba(20, 20, 20, 0.8)';
    styles.backdropFilter = 'blur(12px)';
    styles.borderRadius = '8px';
  }
  
  // Border classes
  if (className.includes('border-gray-900/30')) {
    styles.border = '1px solid rgba(17, 24, 39, 0.3)';
  }
  
  // Add more conversions as needed
  
  return styles;
}

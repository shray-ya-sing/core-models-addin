import * as React from 'react';
import { useState, useEffect } from 'react';

interface TypewriterEffectProps {
  text: string;
  speed?: number;
  loop?: boolean;
}

/**
 * A component that creates a typewriter effect for text
 */
export const TypewriterEffect: React.FC<TypewriterEffectProps> = ({ 
  text, 
  speed = 150, 
  loop = true 
}) => {
  const [displayText, setDisplayText] = useState('');
  const [index, setIndex] = useState(0);
  const [isDeleting, setIsDeleting] = useState(false);
  const [loopCount, setLoopCount] = useState(0);

  useEffect(() => {
    let timer: NodeJS.Timeout;
    
    // If we're at the end of the text and not deleting yet
    if (index === text.length && !isDeleting && loop) {
      // Pause at the end before starting to delete
      timer = setTimeout(() => {
        setIsDeleting(true);
      }, 700);
    } 
    // If we're deleting and reached the beginning
    else if (index === 0 && isDeleting) {
      setIsDeleting(false);
      setLoopCount(loopCount + 1);
    } 
    // Normal typing or deleting
    else {
      timer = setTimeout(() => {
        setIndex(prevIndex => {
          if (isDeleting) {
            return prevIndex - 1;
          } else {
            return prevIndex + 1;
          }
        });
        
        setDisplayText(text.substring(0, isDeleting ? index - 1 : index + 1));
      }, isDeleting ? speed / 2 : speed);
    }
    
    return () => clearTimeout(timer);
  }, [index, isDeleting, text, speed, loop, loopCount]);

  return <span>{displayText}</span>;
};

export default TypewriterEffect;

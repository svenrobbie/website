'use client';

import { motion, useReducedMotion, useScroll, useSpring } from 'motion/react';

export function PageProgress() {
  const reduce = useReducedMotion();
  const { scrollYProgress } = useScroll();
  const scaleX = useSpring(scrollYProgress, { stiffness: 140, damping: 28, mass: 0.25 });

  if (reduce) return null;

  return <motion.div className="page-progress" style={{ scaleX }} aria-hidden="true" />;
}

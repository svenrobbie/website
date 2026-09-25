'use client';

import Image, { type StaticImageData } from 'next/image';
import Link from 'next/link';
import { ArrowRight } from '@phosphor-icons/react';
import {
  motion,
  useMotionValue,
  useReducedMotion,
  useScroll,
  useSpring,
  useTransform,
} from 'motion/react';
import { useRef, type PointerEvent } from 'react';

const ease = [0.16, 1, 0.3, 1] as const;

function MagneticProjectLink() {
  const reduce = useReducedMotion();
  const x = useMotionValue(0);
  const y = useMotionValue(0);
  const springX = useSpring(x, { stiffness: 230, damping: 18, mass: 0.4 });
  const springY = useSpring(y, { stiffness: 230, damping: 18, mass: 0.4 });

  const handleMove = (event: PointerEvent<HTMLAnchorElement>) => {
    if (reduce) return;
    const bounds = event.currentTarget.getBoundingClientRect();
    x.set((event.clientX - bounds.left - bounds.width / 2) * 0.13);
    y.set((event.clientY - bounds.top - bounds.height / 2) * 0.18);
  };

  const reset = () => {
    x.set(0);
    y.set(0);
  };

  return (
    <motion.div style={reduce ? undefined : { x: springX, y: springY }}>
      <Link
        className="button button-primary"
        href="/projects/"
        onPointerMove={handleMove}
        onPointerLeave={reset}
      >
        View projects <ArrowRight size={18} weight="bold" />
      </Link>
    </motion.div>
  );
}

export function Hero({ image }: { image: StaticImageData }) {
  const ref = useRef<HTMLElement>(null);
  const reduce = useReducedMotion();
  const { scrollYProgress } = useScroll({ target: ref, offset: ['start start', 'end start'] });
  const imageY = useTransform(scrollYProgress, [0, 1], ['0%', '4%']);
  const imageScale = useTransform(scrollYProgress, [0, 1], [1.04, 1.01]);
  const copyY = useTransform(scrollYProgress, [0, 1], ['0%', '13%']);
  const imageOpacity = useTransform(scrollYProgress, [0, 0.85], [1, 0.72]);

  return (
    <section className="hero" ref={ref}>
      <motion.div
        className="hero-background"
        initial={reduce ? false : { opacity: 0, scale: 1.025 }}
        animate={{ opacity: 1, scale: 1 }}
        transition={{ duration: 1.15, ease }}
        style={reduce ? undefined : { opacity: imageOpacity }}
      >
        <motion.div className="hero-image-inner" style={reduce ? undefined : { y: imageY, scale: imageScale }}>
          <Image src={image} alt="Panoramic secure infrastructure system with protected compute nodes and illuminated delivery paths" fill priority sizes="100vw" />
        </motion.div>
      </motion.div>

      <div className="shell hero-layout">
        <motion.div
          className="hero-copy"
          style={reduce ? undefined : { y: copyY }}
          initial={reduce ? false : 'hidden'}
          animate="show"
          variants={{
            hidden: {},
            show: { transition: { staggerChildren: 0.09, delayChildren: 0.08 } },
          }}
        >
          <motion.p className="eyebrow" variants={{ hidden: { opacity: 0, y: 14 }, show: { opacity: 1, y: 0, transition: { duration: 0.65, ease } } }}>
            Cybersecurity student / DevOps focus
          </motion.p>
          <motion.h1 variants={{ hidden: { opacity: 0, y: 28 }, show: { opacity: 1, y: 0, transition: { duration: 0.8, ease } } }}>
            I build secure systems that hold up.
          </motion.h1>
          <motion.p className="hero-intro" variants={{ hidden: { opacity: 0, y: 18 }, show: { opacity: 1, y: 0, transition: { duration: 0.7, ease } } }}>
            Security, automation, and infrastructure shaped by real projects and operated beyond the first deploy.
          </motion.p>
          <motion.div className="hero-actions" variants={{ hidden: { opacity: 0, y: 16 }, show: { opacity: 1, y: 0, transition: { duration: 0.65, ease } } }}>
            <MagneticProjectLink />
            <Link className="text-link animated-link" href="/about/">About me</Link>
          </motion.div>
        </motion.div>
      </div>
    </section>
  );
}

'use client';

import Link from 'next/link';
import Image from 'next/image';
import { ArrowUpRight } from '@phosphor-icons/react';
import { motion, useReducedMotion } from 'motion/react';
import type { Project } from '@/data/site';

export function ProjectList({ items }: { items: Project[] }) {
  const reduce = useReducedMotion();

  return (
    <div className="project-list">
      {items.map((project, index) => (
        <motion.div
          key={project.slug}
          initial={reduce ? false : { opacity: 0, y: 24 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true, amount: 0.35 }}
          transition={{ duration: 0.7, delay: index * 0.07, ease: [0.16, 1, 0.3, 1] }}
        >
          <Link
            href={project.href}
            className="project-row"
            target={project.href.startsWith('http') ? '_blank' : undefined}
            rel={project.href.startsWith('http') ? 'noreferrer' : undefined}
          >
            <Image
              className="project-card-image"
              src={project.image}
              alt={project.imageAlt}
              fill
              loading={index === 0 ? 'eager' : 'lazy'}
              sizes="(max-width: 900px) 100vw, 50vw"
            />
            <span className="project-card-scrim" aria-hidden="true" />
            <div className="project-card-content">
              <div>
                <p className="project-area">{project.area}</p>
                <h3>{project.title}</h3>
              </div>
              <p>{project.summary}</p>
              <div className="project-meta">
                <span>{project.status}</span>
                <span>{project.stack.join(' / ')}</span>
              </div>
            </div>
            <span className="project-arrow"><ArrowUpRight size={20} weight="bold" aria-hidden="true" /></span>
          </Link>
        </motion.div>
      ))}
    </div>
  );
}

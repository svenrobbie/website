import type { Metadata } from 'next';
import { ProjectList } from '@/components/ProjectList';
import { projects } from '@/data/site';

export const metadata: Metadata = { title: 'Projects' };

export default function ProjectsPage() {
  return (
    <section className="shell page-intro projects-page">
      <p className="eyebrow">Projects</p>
      <h1>Work shaped by curiosity and real constraints.</h1>
      <p>Original builds and personal forks across security, embedded systems, Android development, and self-hosted infrastructure.</p>
      <ProjectList items={projects} />
    </section>
  );
}

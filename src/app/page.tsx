import Image from 'next/image';
import Link from 'next/link';
import {
  ArrowRight,
  ArrowUpRight,
  ArrowsClockwise,
  BookOpenText,
  ChartLineUp,
  GitBranch,
  GithubLogo,
  HardDrives,
  LinkedinLogo,
  ShieldCheck,
} from '@phosphor-icons/react/dist/ssr';
import { Hero } from '@/components/Hero';
import { Reveal } from '@/components/Reveal';
import { ProjectList } from '@/components/ProjectList';
import { projects, site } from '@/data/site';
import heroIllustration from '../../assest/home-hero-background.webp';
import portrait from '../../assest/sven-pixel-character.webp';

type HomeIcon = typeof ShieldCheck;

type Skill = {
  title: string;
  description: string;
  evidence: string;
  href: string;
  icon: HomeIcon;
  tone: 'violet' | 'neutral' | 'orange';
};

type Principle = {
  title: string;
  description: string;
  icon: HomeIcon;
};

const skills: Skill[] = [
  {
    title: 'Cybersecurity',
    description: 'Vulnerability research, secure delivery, and defensive thinking applied to systems with real trust boundaries.',
    evidence: 'C2 framework',
    href: 'https://github.com/svenrobbie/c2-framework',
    icon: ShieldCheck,
    tone: 'violet',
  },
  {
    title: 'DevOps and automation',
    description: 'Pipelines, repeatable configuration, and deployment workflows designed to make change safer and easier to inspect.',
    evidence: 'Homelab platform',
    href: '/homelab/',
    icon: GitBranch,
    tone: 'neutral',
  },
  {
    title: 'Infrastructure and operations',
    description: 'Linux, containers, observability, resilient networking, backups, and recovery considered as one operating system.',
    evidence: 'MeshCore',
    href: 'https://github.com/svenrobbie/MeshCore',
    icon: HardDrives,
    tone: 'orange',
  },
];

const principles: Principle[] = [
  {
    title: 'Document decisions',
    description: 'Capture intent, dependencies, and recovery steps while the context is still fresh.',
    icon: BookOpenText,
  },
  {
    title: 'Automate repetition',
    description: 'Turn reliable manual work into reviewable configuration and repeatable workflows.',
    icon: GitBranch,
  },
  {
    title: 'Observe behavior',
    description: 'Make health, drift, and failure visible before diagnosis becomes guesswork.',
    icon: ChartLineUp,
  },
  {
    title: 'Plan recovery',
    description: 'Treat backups and restore paths as part of the system, not cleanup after failure.',
    icon: ArrowsClockwise,
  },
];

export default function Home() {
  return (
    <>
      <Hero image={heroIllustration} />

      <section className="section-space home-projects" id="work">
        <div className="shell section-heading">
          <h2>Selected work</h2>
          <Link className="text-link" href="/projects/">All projects <ArrowRight size={16} /></Link>
        </div>
        <ProjectList items={projects.filter((project) => project.featured)} />
      </section>

      <section className="shell section-space skills-section" id="skills">
        <div className="home-section-copy">
          <h2>The skills behind the systems.</h2>
          <p>Technical range matters most when it produces work that can be inspected, operated, and improved.</p>
        </div>
        <div className="skills-grid">
          {skills.map((skill, index) => {
            const Icon = skill.icon;
            const external = skill.href.startsWith('http');

            return (
              <Reveal key={skill.title} delay={index * 0.06}>
                <article className={`skill-card skill-card-${skill.tone}`}>
                  <Icon size={32} weight="duotone" aria-hidden="true" />
                  <div>
                    <h3>{skill.title}</h3>
                    <p>{skill.description}</p>
                  </div>
                  <Link href={skill.href} target={external ? '_blank' : undefined} rel={external ? 'noreferrer' : undefined}>
                    {skill.evidence} <ArrowUpRight size={17} weight="bold" aria-hidden="true" />
                  </Link>
                </article>
              </Reveal>
            );
          })}
        </div>
      </section>

      <section className="shell section-space approach-section" id="approach">
        <Reveal>
          <article className="approach-panel">
            <div className="approach-heading">
              <h2>Built to be understood, operated and improved.</h2>
              <p>The useful part of a system begins after the first successful deployment.</p>
            </div>
            <div className="approach-list">
              {principles.map((principle) => {
                const Icon = principle.icon;

                return (
                  <div className="approach-item" key={principle.title}>
                    <Icon size={24} weight="duotone" aria-hidden="true" />
                    <div><h3>{principle.title}</h3><p>{principle.description}</p></div>
                  </div>
                );
              })}
            </div>
          </article>
        </Reveal>
      </section>

      <section className="shell section-space home-contact-section" id="contact">
        <Reveal>
          <article className="home-contact">
            <div className="home-contact-copy">
              <h2>Want to talk security, infrastructure, or the work behind these projects?</h2>
              <p>I am always interested in practical systems, thoughtful engineering, and opportunities to keep learning.</p>
              <div className="home-contact-links">
                <a href={site.github} target="_blank" rel="noreferrer"><GithubLogo size={23} aria-hidden="true" /> GitHub <ArrowUpRight size={17} weight="bold" aria-hidden="true" /></a>
                <a href={site.linkedin} target="_blank" rel="noreferrer"><LinkedinLogo size={23} aria-hidden="true" /> LinkedIn <ArrowUpRight size={17} weight="bold" aria-hidden="true" /></a>
              </div>
            </div>
            <div className="home-contact-portrait">
              <Image src={portrait} alt="Full-body pixel-art character of Sven van de Lagemaat holding a laptop" fill sizes="(max-width: 900px) 100vw, 44vw" />
            </div>
          </article>
        </Reveal>
      </section>
    </>
  );
}

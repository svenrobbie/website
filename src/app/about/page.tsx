import type { Metadata } from 'next';
import Image from 'next/image';
import { ArrowUpRight } from '@phosphor-icons/react/dist/ssr';
import { site } from '@/data/site';
import portrait from '../../../assest/sven-pixel-portrait.webp';

export const metadata: Metadata = { title: 'About' };

export default function AboutPage() {
  return (
    <section className="shell page-intro about-layout">
      <div className="about-content">
        <p className="eyebrow">About</p>
        <h1>Curious about the whole path from commit to production.</h1>
        <div className="about-copy">
          <p>I am Sven van de Lagemaat, a cybersecurity student and programmer focused on secure systems and DevOps.</p>
          <p>I learn by building: exploring ethical hacking and vulnerability research, automating repeatable work, and operating services at home.</p>
          <p>This portfolio is an evolving record of that work. Project details, technical decisions, and lessons learned will be added as the work develops.</p>
          <a className="button button-primary" href={site.linkedin} target="_blank" rel="noreferrer">
            LinkedIn <ArrowUpRight size={18} weight="bold" aria-hidden="true" />
          </a>
        </div>
      </div>
      <figure className="about-portrait">
        <div className="about-portrait-frame">
          <Image
            src={portrait}
            alt="Pixel-art portrait of Sven van de Lagemaat in a developer workspace"
            priority
            sizes="(max-width: 900px) 100vw, 50vw"
          />
        </div>
        <figcaption><span>01</span> Sven / Developer &amp; security student</figcaption>
      </figure>
    </section>
  );
}

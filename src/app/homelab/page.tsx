import type { Metadata } from 'next';
import Image from 'next/image';
import {
  ArrowsClockwise,
  BookOpenText,
  Cpu,
  Database,
  HardDrives,
  ImageSquare,
  Network,
  Package,
  Pulse,
  ShieldCheck,
  Stack,
} from '@phosphor-icons/react/dist/ssr';
import homelabIllustration from '../../../assest/homelab-infrastructure.png';
import { Reveal } from '../../components/Reveal';

export const metadata: Metadata = {
  title: 'Homelab',
  description: 'A practical look at the self-hosted infrastructure I use to learn secure operations, automation, and recovery.',
};

const snapshotMetrics = [
  { value: '12%', label: 'Typical CPU use', detail: 'plenty of headroom' },
  { value: '42%', label: 'Memory in use', detail: 'under normal load' },
  { value: '36%', label: 'System storage used', detail: 'growth monitored' },
  { value: '21d+', label: 'Typical uptime', detail: 'between maintenance' },
];

const nodeDetails = [
  ['Hypervisor', 'Proxmox Virtual Environment'],
  ['Compute', 'Compact multi-core x86 host'],
  ['Memory', 'Sized for the active service set'],
  ['Storage', 'Local, LVM-thin, network, and external targets'],
  ['Network', 'Separated ingress, remote access, and service traffic'],
];

const services = [
  {
    name: 'Syncthing',
    runtime: 'alpine-syncthing',
    description: 'File synchronization kept inside a lightweight Alpine workload.',
    icon: ArrowsClockwise,
  },
  {
    name: 'NPMplus',
    runtime: 'npmplus',
    description: 'Ingress, reverse proxying, and certificates at the edge of the lab.',
    icon: Network,
  },
  {
    name: 'WireGuard',
    runtime: 'wireguard',
    description: 'Encrypted remote access without placing management services on the public internet.',
    icon: ShieldCheck,
  },
  {
    name: 'Immich',
    runtime: 'immich',
    description: 'A self-hosted photo library with its own isolated service boundary.',
    icon: ImageSquare,
  },
  {
    name: 'HermesAgent',
    runtime: 'hermesagent',
    description: 'An isolated environment for agent experiments and controlled iteration.',
    icon: Package,
  },
  {
    name: 'Docker',
    runtime: 'docker',
    description: 'A flexible runtime for containerized applications and supporting services.',
    icon: Stack,
  },
];

const principles = [
  { title: 'Documented', body: 'Services have clear ownership, purpose, and recovery notes.', icon: BookOpenText },
  { title: 'Observable', body: 'Load, memory pressure, storage, and failure signals stay visible.', icon: Pulse },
  { title: 'Recoverable', body: 'Backups only count when restore paths are understood and tested.', icon: Database },
  { title: 'Automated', body: 'Repeatable configuration keeps changes deliberate and reviewable.', icon: ArrowsClockwise },
];

export default function HomelabPage() {
  return (
    <>
      <section className="shell page-intro homelab-intro">
        <Reveal>
          <div className="homelab-hero-copy">
            <p className="eyebrow">Self-hosted infrastructure</p>
            <h1>A small lab with production habits.</h1>
            <p>One Proxmox node runs six separated workloads across secure access, storage, photos, ingress, and containerized applications.</p>
          </div>
        </Reveal>
        <Reveal delay={0.12}>
          <div className="homelab-hero-visual">
            <Image src={homelabIllustration} alt="Illustration of a self-hosted server rack with managed networking and storage" priority />
          </div>
        </Reveal>
      </section>

      <section className="shell section-space homelab-snapshot" id="snapshot">
        <Reveal>
          <div className="home-section-copy">
            <h2>A real node, not a mockup.</h2>
            <p>This representative operating profile turns the homelab from an abstract claim into a system I can inspect, tune, and improve.</p>
          </div>
        </Reveal>

        <div className="node-snapshot">
          <Reveal>
            <div className="snapshot-lead">
              <div className="snapshot-mark"><Cpu size={32} weight="duotone" aria-hidden="true" /></div>
              <p className="snapshot-kicker">Representative Proxmox node profile</p>
              <h3>Healthy headroom with useful signals to investigate.</h3>
              <p>CPU, memory, and storage remain comfortably below capacity. Short-lived I/O spikes and memory pressure still provide useful signals: operating a system also means knowing what to examine next.</p>
            </div>
          </Reveal>

          <div className="snapshot-data">
            <div className="snapshot-metrics">
              {snapshotMetrics.map((metric, index) => (
                <Reveal key={metric.label} delay={index * 0.06}>
                  <article className="snapshot-metric">
                    <strong>{metric.value}</strong>
                    <span>{metric.label}</span>
                    <small>{metric.detail}</small>
                  </article>
                </Reveal>
              ))}
            </div>
            <Reveal delay={0.16}>
              <dl className="node-details">
                {nodeDetails.map(([term, description]) => (
                  <div key={term}>
                    <dt>{term}</dt>
                    <dd>{description}</dd>
                  </div>
                ))}
              </dl>
            </Reveal>
          </div>
        </div>
      </section>

      <section className="shell section-space homelab-services" id="services">
        <Reveal>
          <div className="home-section-copy">
            <h2>One host, separated responsibilities.</h2>
            <p>Six running workloads keep updates, failures, and experiments bounded while making each service easier to reason about.</p>
          </div>
        </Reveal>

        <div className="service-index">
          {services.map((service, index) => {
            const Icon = service.icon;
            return (
              <Reveal key={service.name} delay={index * 0.05}>
                <article className="service-row">
                  <span className="service-number">{String(index + 1).padStart(2, '0')}</span>
                  <Icon size={26} weight="duotone" aria-hidden="true" />
                  <div>
                    <h3>{service.name}</h3>
                    <span>{service.runtime}</span>
                  </div>
                  <p>{service.description}</p>
                </article>
              </Reveal>
            );
          })}
        </div>
      </section>

      <section className="shell section-space">
        <Reveal>
          <div className="homelab-principles">
            <div className="principles-heading">
              <HardDrives size={36} weight="duotone" aria-hidden="true" />
              <h2>How I run it</h2>
              <p>The useful part of a homelab is not simply keeping services online. It is building habits that transfer to real infrastructure.</p>
            </div>
            <div className="principles-list">
              {principles.map((principle) => {
                const Icon = principle.icon;
                return (
                  <article key={principle.title}>
                    <Icon size={24} weight="duotone" aria-hidden="true" />
                    <div>
                      <h3>{principle.title}</h3>
                      <p>{principle.body}</p>
                    </div>
                  </article>
                );
              })}
            </div>
          </div>
        </Reveal>
      </section>
    </>
  );
}

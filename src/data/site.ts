import type { StaticImageData } from 'next/image';
import c2FrameworkImage from '../../assest/c2-framework.webp';
import flip2DndImage from '../../assest/flip-2-dnd.webp';
import homelabImage from '../../assest/homelab-project.webp';
import meshCoreImage from '../../assest/meshcore.webp';
import moneyWiseImage from '../../assest/moneywise.webp';

export const site = {
  name: 'Sven van de Lagemaat',
  shortName: 'RogueByte',
  handle: 'RogueByte',
  role: 'Cybersecurity student and programmer focused on DevOps',
  github: 'https://github.com/svenrobbie',
  linkedin: 'https://www.linkedin.com/in/svenvdla/',
};

export type Project = {
  slug: string;
  title: string;
  summary: string;
  area: string;
  status: string;
  stack: string[];
  href: string;
  image: StaticImageData;
  imageAlt: string;
  featured?: boolean;
};

export const projects: Project[] = [
  {
    slug: 'c2-framework',
    title: 'C2 framework',
    summary: 'A Python command-and-control framework with a C2 server and front-end.',
    area: 'Cybersecurity',
    status: 'Public repository',
    stack: ['Python', 'C2', 'Security'],
    href: 'https://github.com/svenrobbie/c2-framework',
    image: c2FrameworkImage,
    imageAlt: 'Abstract command node coordinating isolated network endpoints',
    featured: true,
  },
  {
    slug: 'meshcore',
    title: 'MeshCore',
    summary: 'A personal C++ fork adding custom M5Stack interfaces and a shared-SPI display driver to a resilient LoRa mesh.',
    area: 'Embedded systems',
    status: 'Personal fork',
    stack: ['C++', 'LoRa', 'M5Stack'],
    href: 'https://github.com/svenrobbie/MeshCore',
    image: meshCoreImage,
    imageAlt: 'Embedded packet-radio modules communicating across a multi-hop mesh',
    featured: true,
  },
  {
    slug: 'homelab-platform',
    title: 'Homelab platform',
    summary: 'Self-hosted services treated as production infrastructure: observable, documented, and recoverable.',
    area: 'Infrastructure',
    status: 'Active',
    stack: ['Linux', 'Ansible', 'Containers'],
    href: '/homelab/',
    image: homelabImage,
    imageAlt: 'Modular self-hosted infrastructure rack with connected services',
    featured: true,
  },
  {
    slug: 'moneywise',
    title: 'MoneyWise',
    summary: 'A Kotlin and Jetpack Compose app for salary calculations, work-time cost, savings and portfolio planning.',
    area: 'Application development',
    status: 'Public repository',
    stack: ['Kotlin', 'Compose', 'Material 3'],
    href: 'https://github.com/svenrobbie/MoneyWise',
    image: moneyWiseImage,
    imageAlt: 'Sculptural mobile finance dashboard with planning and allocation forms',
  },
  {
    slug: 'flip-2-dnd',
    title: 'Flip 2 DND',
    summary: 'An open Android fork with the paywall removed, improved battery behavior, Hilt injection and modernized tooling.',
    area: 'Android utility',
    status: 'Personal fork',
    stack: ['Kotlin', 'Compose', 'Hilt'],
    href: 'https://github.com/svenrobbie/flip_2_dnd',
    image: flip2DndImage,
    imageAlt: 'Smartphone rotating face down to activate a quiet focus mode',
  },
];

import Link from 'next/link';
import {
  ArrowUpRight,
  FolderOpen,
  HardDrives,
  LinkedinLogo,
  ShieldChevron,
  User,
} from '@phosphor-icons/react/dist/ssr';
import { site } from '@/data/site';

const navigation = [
  { href: '/projects/', label: 'Projects', icon: FolderOpen },
  { href: '/homelab/', label: 'Homelab', icon: HardDrives },
  { href: '/about/', label: 'About', icon: User },
];

export function Header() {
  return (
    <header className="site-header">
      <div className="shell nav-wrap">
        <Link className="brand" href="/" aria-label={`${site.name}, home`}>
          <span className="brand-symbol" aria-hidden="true">
            <ShieldChevron size={23} weight="duotone" />
          </span>
          <span>{site.shortName}</span>
        </Link>
        <nav aria-label="Primary navigation">
          {navigation.map((item) => {
            const Icon = item.icon;
            return <Link key={item.href} href={item.href}><Icon size={16} weight="duotone" aria-hidden="true" />{item.label}</Link>;
          })}
        </nav>
        <a className="nav-contact" href={site.linkedin} target="_blank" rel="noreferrer">
          <LinkedinLogo size={17} weight="fill" aria-hidden="true" />
          LinkedIn
          <ArrowUpRight size={15} weight="bold" aria-hidden="true" />
        </a>
      </div>
    </header>
  );
}

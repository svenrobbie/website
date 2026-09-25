import { GithubLogo, LinkedinLogo } from '@phosphor-icons/react/dist/ssr';
import { site } from '@/data/site';

export function Footer() {
  return (
    <footer className="footer">
      <div className="shell footer-grid">
        <p><strong>{site.name}</strong><br />{site.handle} on GitHub</p>
        <p className="footer-note">Building dependable systems, then documenting what made them work.</p>
        <div className="social-links">
          <a href={site.github} aria-label="GitHub"><GithubLogo size={22} weight="regular" /></a>
          <a href={site.linkedin} aria-label="LinkedIn"><LinkedinLogo size={22} weight="regular" /></a>
        </div>
      </div>
    </footer>
  );
}

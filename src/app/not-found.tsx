import Link from 'next/link';

export default function NotFound() {
  return <section className="shell page-intro"><p className="eyebrow">404</p><h1>This route is not documented.</h1><p>The page may have moved, or it never existed.</p><Link className="button button-primary" href="/">Back home</Link></section>;
}

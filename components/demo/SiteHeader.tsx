import Link from "next/link";

const LINKS = [
  { href: "/", label: "Demo" },
  { href: "/harness", label: "Iframe harness" },
] as const;

type Props = {
  current: (typeof LINKS)[number]["href"];
};

export function SiteHeader({ current }: Props) {
  return (
    <header className="flex shrink-0 items-center gap-4 border-b border-zinc-800 bg-zinc-950 px-4 py-2">
      <p className="text-sm font-semibold text-zinc-100">Transcript Editor</p>
      <nav className="flex items-center gap-1 text-sm">
        {LINKS.map((link) => (
          <Link
            key={link.href}
            href={link.href}
            className={`rounded-md px-2.5 py-1 ${
              link.href === current
                ? "bg-zinc-800 text-zinc-100"
                : "text-zinc-400 hover:bg-zinc-900 hover:text-zinc-100"
            }`}
          >
            {link.label}
          </Link>
        ))}
        <a
          href="/embed"
          target="_blank"
          rel="noreferrer"
          className="rounded-md px-2.5 py-1 text-zinc-400 hover:bg-zinc-900 hover:text-zinc-100"
        >
          /embed
        </a>
      </nav>
    </header>
  );
}

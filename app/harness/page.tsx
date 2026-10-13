import type { Metadata } from "next";
import { SiteHeader } from "@/components/demo/SiteHeader";
import { HarnessHost } from "@/components/harness/HarnessHost";

export const metadata: Metadata = {
  title: "Transcript Editor (iframe harness)",
};

export default function HarnessPage() {
  return (
    <div className="flex h-screen flex-col">
      <SiteHeader current="/harness" />
      <HarnessHost />
    </div>
  );
}

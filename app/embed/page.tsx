import type { Metadata } from "next";
import { EmbedBridge } from "@/components/embed/EmbedBridge";

export const metadata: Metadata = {
  title: "Transcript Editor (embed)",
  robots: { index: false },
};

export default function EmbedPage() {
  return <EmbedBridge />;
}

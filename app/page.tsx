import { EditorDemo } from "@/components/demo/EditorDemo";
import { SiteHeader } from "@/components/demo/SiteHeader";

export default function DemoPage() {
  return (
    <div className="flex h-screen flex-col">
      <SiteHeader current="/" />
      <EditorDemo />
    </div>
  );
}

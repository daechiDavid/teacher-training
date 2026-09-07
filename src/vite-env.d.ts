/// <reference types="vite/client" />

declare module "*.ttf" {
  const src: string;
  export default src;
}

declare module "lucide-react/dist/esm/icons/*.js" {
  import type { ForwardRefExoticComponent, RefAttributes } from "react";
  import type { LucideProps } from "lucide-react";

  const icon: ForwardRefExoticComponent<
    Omit<LucideProps, "ref"> & RefAttributes<SVGSVGElement>
  >;
  export default icon;
}

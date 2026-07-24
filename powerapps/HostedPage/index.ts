import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class HostedPage
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private frame!: HTMLIFrameElement;
  private currentUrl = "";

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    container.style.width = "100%";
    container.style.height = "100%";
    container.style.overflow = "hidden";

    this.frame = document.createElement("iframe");
    this.frame.style.width = "100%";
    this.frame.style.height = "100%";
    this.frame.style.border = "0";
    this.frame.style.display = "block";
    this.frame.setAttribute(
      "allow",
      "clipboard-read; clipboard-write; fullscreen"
    );

    container.appendChild(this.frame);
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    const nextUrl = context.parameters.siteUrl.raw?.trim() ?? "";

    if (nextUrl !== this.currentUrl) {
      this.currentUrl = nextUrl;
      this.frame.src = nextUrl;
    }
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    this.frame?.remove();
  }
}
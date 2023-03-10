import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/MS600ConfigTab/index.html")
@PreventIframe("/MS600ConfigTab/config.html")
@PreventIframe("/MS600ConfigTab/remove.html")
export class MS600ConfigTab {
}

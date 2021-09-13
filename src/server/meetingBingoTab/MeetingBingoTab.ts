import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/meetingBingoTab/index.html")
@PreventIframe("/meetingBingoTab/config.html")
@PreventIframe("/meetingBingoTab/remove.html")
export class MeetingBingoTab {
}

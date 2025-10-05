import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface ITeamsAiMeetingAppWebPartProps {
    description: string;
}
export default class TeamsAiMeetingAppWebPart extends BaseClientSideWebPart<ITeamsAiMeetingAppWebPartProps> {
    private meetingId;
    private teamName;
    private isInMeeting;
    render(): void;
    private renderMeetingContent;
    private renderWelcomeContent;
    private addEventListeners;
    private getTranscriptionAndSummarize;
    private getMeetingTranscription;
    private parseVTTTranscript;
    private parseJSONTranscript;
    private sendToAIService;
    private generateMockSummary;
    private formatAISummary;
    private delay;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}

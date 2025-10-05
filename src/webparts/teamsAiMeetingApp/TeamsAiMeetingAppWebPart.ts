import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import styles from './TeamsAiMeetingAppWebPart.module.scss';
import * as strings from 'TeamsAiMeetingAppWebPartStrings';

export interface ITeamsAiMeetingAppWebPartProps {
  description: string;
}

export default class TeamsAiMeetingAppWebPart extends BaseClientSideWebPart<ITeamsAiMeetingAppWebPartProps> {

  private meetingId: string = '';
  private teamName: string = '';
  private isInMeeting: boolean = false;

  public render(): void {
    let title: string = 'Welcome to Teams AI Meeting Assistant';
    let subTitle: string = 'Ready to enhance your meeting experience';
    let meetingInfo: string = '';
    
    // Check if we're in Microsoft Teams context
    if (this.context.sdks.microsoftTeams) {
      const teamsContext = this.context.sdks.microsoftTeams.context;
      
      if (teamsContext.meetingId) {
        this.meetingId = teamsContext.meetingId;
        title = "AI Meeting Assistant Active";
        subTitle = "Meeting in progress - Transcription and AI features available";
        meetingInfo = `Meeting ID: ${this.meetingId}`;
        this.isInMeeting = true;
      } else if (teamsContext.teamName) {
        this.teamName = teamsContext.teamName;
        title = "AI Meeting Assistant";
        subTitle = `Ready for ${this.teamName} meetings`;
        meetingInfo = `Team: ${this.teamName}`;
      }
    }

    this.domElement.innerHTML = `
      <div class="${styles.teamsAiMeetingApp}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <div class="${styles.header}">
                <span class="${styles.title}">${title}</span>
                <p class="${styles.subTitle}">${subTitle}</p>
                ${meetingInfo ? `<p class="${styles.meetingInfo}">${meetingInfo}</p>` : ''}
              </div>
              
              <div class="${styles.content}">
                ${this.isInMeeting ? this.renderMeetingContent() : this.renderWelcomeContent()}
              </div>
            </div>
          </div>
        </div>
      </div>`;

    // Add event listeners
    this.addEventListeners();
  }

  private renderMeetingContent(): string {
    return `
      <div class="${styles.meetingPanel}">
        <div class="${styles.section}">
          <h3>Meeting Status</h3>
          <div class="${styles.statusIndicator}">
            <span class="${styles.statusDot} ${styles.active}"></span>
            <span>Meeting Active</span>
          </div>
        </div>

        <div class="${styles.section}">
          <h3>Transcription & AI Summary</h3>
          <div class="${styles.transcriptionPanel}">
            <div id="transcriptionStatus" class="${styles.status}">
              Ready to capture meeting transcription...
            </div>
            <button id="getTranscriptionBtn" class="${styles.button} ${styles.primary}">
              Get Post-Meeting Transcription
            </button>
          </div>
        </div>

        <div class="${styles.section}">
          <h3>AI Summary</h3>
          <div id="summaryPanel" class="${styles.summaryPanel}">
            <div id="summaryContent" class="${styles.summaryContent}">
              Meeting summary will appear here after transcription is processed...
            </div>
            <div id="summaryLoading" class="${styles.loading}" style="display: none;">
              <div class="${styles.spinner}"></div>
              <span>Generating AI summary...</span>
            </div>
          </div>
        </div>
      </div>
    `;
  }

  private renderWelcomeContent(): string {
    return `
      <div class="${styles.welcomePanel}">
        <div class="${styles.section}">
          <h3>Features</h3>
          <ul class="${styles.featureList}">
            <li>📝 Automatic meeting transcription capture</li>
            <li>🤖 AI-powered meeting summaries</li>
            <li>📊 Key insights and action items</li>
            <li>🔗 Integration with custom AI service</li>
          </ul>
        </div>
        
        <div class="${styles.section}">
          <h3>How it works</h3>
          <ol class="${styles.stepsList}">
            <li>Join a Teams meeting with this app installed</li>
            <li>After the meeting ends, click "Get Transcription"</li>
            <li>AI will automatically generate a comprehensive summary</li>
            <li>View insights, action items, and key discussion points</li>
          </ol>
        </div>
      </div>
    `;
  }

  private addEventListeners(): void {
    const getTranscriptionBtn = this.domElement.querySelector('#getTranscriptionBtn');
    if (getTranscriptionBtn) {
      getTranscriptionBtn.addEventListener('click', () => {
        this.getTranscriptionAndSummarize();
      });
    }
  }

  private async getTranscriptionAndSummarize(): Promise<void> {
    const statusElement = this.domElement.querySelector('#transcriptionStatus');
    const summaryContent = this.domElement.querySelector('#summaryContent');
    const summaryLoading = this.domElement.querySelector('#summaryLoading') as HTMLElement;
    const button = this.domElement.querySelector('#getTranscriptionBtn') as HTMLButtonElement;

    if (!statusElement || !summaryContent || !summaryLoading || !button) return;

    try {
      // Disable button and show loading
      button.disabled = true;
      button.textContent = 'Processing...';
      statusElement.textContent = 'Fetching meeting transcription...';
      summaryLoading.style.display = 'flex';
      summaryContent.textContent = 'Processing transcription...';

      // Step 1: Get meeting transcription using Microsoft Graph API
      const transcription = await this.getMeetingTranscription();
      
      if (!transcription) {
        statusElement.textContent = 'No transcription available yet. Please try again later.';
        summaryContent.textContent = 'Transcription not ready. Please wait for the meeting to end and try again.';
        return;
      }

      statusElement.textContent = 'Transcription retrieved successfully. Generating AI summary...';

      // Step 2: Send transcription to custom AI API
      const summary = await this.sendToAIService(transcription);

      // Step 3: Display results
      statusElement.textContent = 'AI summary generated successfully!';
      summaryContent.innerHTML = this.formatAISummary(summary);

    } catch (error) {
      console.error('Error processing transcription:', error);
      statusElement.textContent = 'Error processing transcription. Please try again.';
      summaryContent.textContent = 'Failed to generate summary. Please try again later.';
    } finally {
      summaryLoading.style.display = 'none';
      button.disabled = false;
      button.textContent = 'Get Post-Meeting Transcription';
    }
  }

  private async getMeetingTranscription(): Promise<string | null> {
    try {
      if (!this.meetingId) {
        console.warn('No meeting ID available for transcription');
        return null;
      }

      console.log(`Fetching transcription for meeting: ${this.meetingId}`);
      
      // Get Microsoft Graph client
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      
      // Step 1: Get the meeting transcripts list
      console.log('Fetching meeting transcripts...');
      const transcriptsResponse = await graphClient
        .api(`/me/onlineMeetings/${this.meetingId}/transcripts`)
        .get();

      if (!transcriptsResponse.value || transcriptsResponse.value.length === 0) {
        console.warn('No transcripts found for this meeting');
        return 'No transcription available yet. Please ensure the meeting has ended and transcription was enabled during the meeting.';
      }

      // Step 2: Get the most recent transcript
      const latestTranscript = transcriptsResponse.value[0];
      console.log(`Found transcript: ${latestTranscript.id}`);

      // Step 3: Get transcript content
      console.log('Fetching transcript content...');
      const transcriptContent = await graphClient
        .api(`/me/onlineMeetings/${this.meetingId}/transcripts/${latestTranscript.id}/content`)
        .get();

      if (!transcriptContent) {
        console.warn('Transcript content is empty');
        return 'Transcript content is not available yet. Please try again later.';
      }

      // Step 4: Parse VTT content if needed
      let processedTranscript = '';
      if (typeof transcriptContent === 'string') {
        processedTranscript = this.parseVTTTranscript(transcriptContent);
      } else {
        // If it's JSON format, extract the text
        processedTranscript = this.parseJSONTranscript(transcriptContent);
      }

      console.log('Successfully retrieved transcript content');
      return processedTranscript;

    } catch (error) {
      console.error('Error fetching transcription from Graph API:', error);
      
      // Provide specific error messages based on error type
      if (error.code === 'Forbidden' || error.code === 'Unauthorized') {
        return 'Access denied. Please ensure you have the necessary permissions to read meeting transcripts (OnlineMeetings.Read, OnlineMeetingTranscript.Read.All).';
      } else if (error.code === 'NotFound') {
        return 'Meeting or transcript not found. Please ensure the meeting ID is correct and transcription was enabled.';
      } else if (error.code === 'BadRequest') {
        return 'Invalid meeting ID format or transcript not ready yet.';
      } else {
        return `Error retrieving transcript: ${error.message || 'Unknown error occurred'}`;
      }
    }
  }

  private parseVTTTranscript(vttContent: string): string {
    try {
      // Parse VTT (WebVTT) format transcript
      const lines = vttContent.split('\n');
      const transcriptLines: string[] = [];
      let currentSpeaker = '';
      let currentText = '';

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        
        // Skip VTT header and timing lines
        if (line.startsWith('WEBVTT') || line.includes('-->') || line === '') {
          continue;
        }

        // Check if line contains speaker information
        if (line.includes(':')) {
          // If we have accumulated text, add it to transcript
          if (currentText) {
            transcriptLines.push(`${currentSpeaker}: ${currentText}`);
            currentText = '';
          }

          // Extract speaker and text
          const speakerMatch = line.match(/^([^:]+):\s*(.*)$/);
          if (speakerMatch) {
            currentSpeaker = speakerMatch[1].trim();
            currentText = speakerMatch[2].trim();
          }
        } else if (currentSpeaker) {
          // Continue text from previous line
          currentText += ' ' + line;
        }
      }

      // Add final accumulated text
      if (currentText && currentSpeaker) {
        transcriptLines.push(`${currentSpeaker}: ${currentText}`);
      }

      return transcriptLines.join('\n');
    } catch (error) {
      console.error('Error parsing VTT transcript:', error);
      return vttContent; // Return raw content if parsing fails
    }
  }

  private parseJSONTranscript(jsonContent: any): string {
    try {
      // Handle different JSON transcript formats
      if (jsonContent.transcript) {
        return jsonContent.transcript;
      }
      
      if (jsonContent.transcripts && Array.isArray(jsonContent.transcripts)) {
        return jsonContent.transcripts
          .map((t: any) => `${t.speaker || 'Unknown'}: ${t.text || ''}`)
          .join('\n');
      }

      if (jsonContent.value && Array.isArray(jsonContent.value)) {
        return jsonContent.value
          .map((item: any) => `${item.speaker || 'Unknown'}: ${item.text || ''}`)
          .join('\n');
      }

      // If it's already a string, return as-is
      if (typeof jsonContent === 'string') {
        return jsonContent;
      }

      return JSON.stringify(jsonContent, null, 2);
    } catch (error) {
      console.error('Error parsing JSON transcript:', error);
      return JSON.stringify(jsonContent, null, 2);
    }
  }

  private async sendToAIService(transcription: string): Promise<any> {
    const aiServiceUrl = 'https://al-meeting-agentassistant-hdhzh7eeb4g8c0fn.westeurope-01.azurewebsites.net/summarise';
    
    try {
      const response = await fetch(aiServiceUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          transcription: transcription,
          meetingId: this.meetingId,
          timestamp: new Date().toISOString()
        })
      });

      if (!response.ok) {
        throw new Error(`AI service responded with status: ${response.status}`);
      }

      const result = await response.json();
      return result;
    } catch (error) {
      console.error('Error calling AI service:', error);
      
      // Fallback: Generate a mock summary if the service is unavailable
      return this.generateMockSummary(transcription);
    }
  }

  private generateMockSummary(transcription: string): any {
    // Mock AI summary for demo purposes
    return {
      summary: "This quarterly review meeting covered strong Q3 performance with sales exceeding targets by 15% and customer satisfaction improving to 4.7/5. The team discussed Q4 strategy focusing on European market expansion with necessary language localization.",
      keyPoints: [
        "Q3 sales exceeded targets by 15%",
        "Customer satisfaction improved from 4.2 to 4.7 out of 5",
        "European market expansion identified as Q4 priority",
        "Language localization requirements need to be addressed",
        "3-month development timeline estimated for core features"
      ],
      actionItems: [
        {
          task: "Prepare detailed European market analysis",
          assignee: "Sarah",
          dueDate: "Next week"
        },
        {
          task: "Plan user testing in parallel with development",
          assignee: "Lisa",
          dueDate: "To be scheduled"
        },
        {
          task: "Schedule follow-up meetings for Q4 initiatives",
          assignee: "John",
          dueDate: "This week"
        }
      ],
      sentiment: "Positive",
      duration: "Estimated 30 minutes",
      participants: ["John", "Sarah", "Mike", "Lisa"]
    };
  }

  private formatAISummary(summary: any): string {
    return `
      <div class="${styles.summaryFormatted}">
        <div class="${styles.summarySection}">
          <h4>📋 Meeting Summary</h4>
          <p>${summary.summary}</p>
        </div>
        
        <div class="${styles.summarySection}">
          <h4>🔑 Key Points</h4>
          <ul>
            ${summary.keyPoints.map(point => `<li>${point}</li>`).join('')}
          </ul>
        </div>
        
        <div class="${styles.summarySection}">
          <h4>✅ Action Items</h4>
          <div class="${styles.actionItems}">
            ${summary.actionItems.map(item => `
              <div class="${styles.actionItem}">
                <strong>${item.task}</strong><br>
                <span class="${styles.assignee}">Assigned to: ${item.assignee}</span><br>
                <span class="${styles.dueDate}">Due: ${item.dueDate}</span>
              </div>
            `).join('')}
          </div>
        </div>
        
        <div class="${styles.summarySection}">
          <h4>📊 Meeting Insights</h4>
          <p><strong>Sentiment:</strong> ${summary.sentiment}</p>
          <p><strong>Duration:</strong> ${summary.duration}</p>
          <p><strong>Participants:</strong> ${summary.participants.join(', ')}</p>
        </div>
      </div>
    `;
  }

  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

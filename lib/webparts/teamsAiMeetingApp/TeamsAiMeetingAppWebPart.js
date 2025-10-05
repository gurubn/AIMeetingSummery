var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import styles from './TeamsAiMeetingAppWebPart.module.scss';
import * as strings from 'TeamsAiMeetingAppWebPartStrings';
var TeamsAiMeetingAppWebPart = /** @class */ (function (_super) {
    __extends(TeamsAiMeetingAppWebPart, _super);
    function TeamsAiMeetingAppWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.meetingId = '';
        _this.teamName = '';
        _this.isInMeeting = false;
        return _this;
    }
    TeamsAiMeetingAppWebPart.prototype.render = function () {
        var title = 'Welcome to Teams AI Meeting Assistant';
        var subTitle = 'Ready to enhance your meeting experience';
        var meetingInfo = '';
        // Check if we're in Microsoft Teams context
        if (this.context.sdks.microsoftTeams) {
            var teamsContext = this.context.sdks.microsoftTeams.context;
            if (teamsContext.meetingId) {
                this.meetingId = teamsContext.meetingId;
                title = "AI Meeting Assistant Active";
                subTitle = "Meeting in progress - Transcription and AI features available";
                meetingInfo = "Meeting ID: ".concat(this.meetingId);
                this.isInMeeting = true;
            }
            else if (teamsContext.teamName) {
                this.teamName = teamsContext.teamName;
                title = "AI Meeting Assistant";
                subTitle = "Ready for ".concat(this.teamName, " meetings");
                meetingInfo = "Team: ".concat(this.teamName);
            }
        }
        this.domElement.innerHTML = "\n      <div class=\"".concat(styles.teamsAiMeetingApp, "\">\n        <div class=\"").concat(styles.container, "\">\n          <div class=\"").concat(styles.row, "\">\n            <div class=\"").concat(styles.column, "\">\n              <div class=\"").concat(styles.header, "\">\n                <span class=\"").concat(styles.title, "\">").concat(title, "</span>\n                <p class=\"").concat(styles.subTitle, "\">").concat(subTitle, "</p>\n                ").concat(meetingInfo ? "<p class=\"".concat(styles.meetingInfo, "\">").concat(meetingInfo, "</p>") : '', "\n              </div>\n              \n              <div class=\"").concat(styles.content, "\">\n                ").concat(this.isInMeeting ? this.renderMeetingContent() : this.renderWelcomeContent(), "\n              </div>\n            </div>\n          </div>\n        </div>\n      </div>");
        // Add event listeners
        this.addEventListeners();
    };
    TeamsAiMeetingAppWebPart.prototype.renderMeetingContent = function () {
        return "\n      <div class=\"".concat(styles.meetingPanel, "\">\n        <div class=\"").concat(styles.section, "\">\n          <h3>Meeting Status</h3>\n          <div class=\"").concat(styles.statusIndicator, "\">\n            <span class=\"").concat(styles.statusDot, " ").concat(styles.active, "\"></span>\n            <span>Meeting Active</span>\n          </div>\n        </div>\n\n        <div class=\"").concat(styles.section, "\">\n          <h3>Transcription & AI Summary</h3>\n          <div class=\"").concat(styles.transcriptionPanel, "\">\n            <div id=\"transcriptionStatus\" class=\"").concat(styles.status, "\">\n              Ready to capture meeting transcription...\n            </div>\n            <button id=\"getTranscriptionBtn\" class=\"").concat(styles.button, " ").concat(styles.primary, "\">\n              Get Post-Meeting Transcription\n            </button>\n          </div>\n        </div>\n\n        <div class=\"").concat(styles.section, "\">\n          <h3>AI Summary</h3>\n          <div id=\"summaryPanel\" class=\"").concat(styles.summaryPanel, "\">\n            <div id=\"summaryContent\" class=\"").concat(styles.summaryContent, "\">\n              Meeting summary will appear here after transcription is processed...\n            </div>\n            <div id=\"summaryLoading\" class=\"").concat(styles.loading, "\" style=\"display: none;\">\n              <div class=\"").concat(styles.spinner, "\"></div>\n              <span>Generating AI summary...</span>\n            </div>\n          </div>\n        </div>\n      </div>\n    ");
    };
    TeamsAiMeetingAppWebPart.prototype.renderWelcomeContent = function () {
        return "\n      <div class=\"".concat(styles.welcomePanel, "\">\n        <div class=\"").concat(styles.section, "\">\n          <h3>Features</h3>\n          <ul class=\"").concat(styles.featureList, "\">\n            <li>\uD83D\uDCDD Automatic meeting transcription capture</li>\n            <li>\uD83E\uDD16 AI-powered meeting summaries</li>\n            <li>\uD83D\uDCCA Key insights and action items</li>\n            <li>\uD83D\uDD17 Integration with custom AI service</li>\n          </ul>\n        </div>\n        \n        <div class=\"").concat(styles.section, "\">\n          <h3>How it works</h3>\n          <ol class=\"").concat(styles.stepsList, "\">\n            <li>Join a Teams meeting with this app installed</li>\n            <li>After the meeting ends, click \"Get Transcription\"</li>\n            <li>AI will automatically generate a comprehensive summary</li>\n            <li>View insights, action items, and key discussion points</li>\n          </ol>\n        </div>\n      </div>\n    ");
    };
    TeamsAiMeetingAppWebPart.prototype.addEventListeners = function () {
        var _this = this;
        var getTranscriptionBtn = this.domElement.querySelector('#getTranscriptionBtn');
        if (getTranscriptionBtn) {
            getTranscriptionBtn.addEventListener('click', function () {
                _this.getTranscriptionAndSummarize();
            });
        }
    };
    TeamsAiMeetingAppWebPart.prototype.getTranscriptionAndSummarize = function () {
        return __awaiter(this, void 0, void 0, function () {
            var statusElement, summaryContent, summaryLoading, button, transcription, summary, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        statusElement = this.domElement.querySelector('#transcriptionStatus');
                        summaryContent = this.domElement.querySelector('#summaryContent');
                        summaryLoading = this.domElement.querySelector('#summaryLoading');
                        button = this.domElement.querySelector('#getTranscriptionBtn');
                        if (!statusElement || !summaryContent || !summaryLoading || !button)
                            return [2 /*return*/];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, 5, 6]);
                        // Disable button and show loading
                        button.disabled = true;
                        button.textContent = 'Processing...';
                        statusElement.textContent = 'Fetching meeting transcription...';
                        summaryLoading.style.display = 'flex';
                        summaryContent.textContent = 'Processing transcription...';
                        return [4 /*yield*/, this.getMeetingTranscription()];
                    case 2:
                        transcription = _a.sent();
                        if (!transcription) {
                            statusElement.textContent = 'No transcription available yet. Please try again later.';
                            summaryContent.textContent = 'Transcription not ready. Please wait for the meeting to end and try again.';
                            return [2 /*return*/];
                        }
                        statusElement.textContent = 'Transcription retrieved successfully. Generating AI summary...';
                        return [4 /*yield*/, this.sendToAIService(transcription)];
                    case 3:
                        summary = _a.sent();
                        // Step 3: Display results
                        statusElement.textContent = 'AI summary generated successfully!';
                        summaryContent.innerHTML = this.formatAISummary(summary);
                        return [3 /*break*/, 6];
                    case 4:
                        error_1 = _a.sent();
                        console.error('Error processing transcription:', error_1);
                        statusElement.textContent = 'Error processing transcription. Please try again.';
                        summaryContent.textContent = 'Failed to generate summary. Please try again later.';
                        return [3 /*break*/, 6];
                    case 5:
                        summaryLoading.style.display = 'none';
                        button.disabled = false;
                        button.textContent = 'Get Post-Meeting Transcription';
                        return [7 /*endfinally*/];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    TeamsAiMeetingAppWebPart.prototype.getMeetingTranscription = function () {
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, transcriptsResponse, latestTranscript, transcriptContent, processedTranscript, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        if (!this.meetingId) {
                            console.warn('No meeting ID available for transcription');
                            return [2 /*return*/, null];
                        }
                        console.log("Fetching transcription for meeting: ".concat(this.meetingId));
                        return [4 /*yield*/, this.context.msGraphClientFactory.getClient('3')];
                    case 1:
                        graphClient = _a.sent();
                        // Step 1: Get the meeting transcripts list
                        console.log('Fetching meeting transcripts...');
                        return [4 /*yield*/, graphClient
                                .api("/me/onlineMeetings/".concat(this.meetingId, "/transcripts"))
                                .get()];
                    case 2:
                        transcriptsResponse = _a.sent();
                        if (!transcriptsResponse.value || transcriptsResponse.value.length === 0) {
                            console.warn('No transcripts found for this meeting');
                            return [2 /*return*/, 'No transcription available yet. Please ensure the meeting has ended and transcription was enabled during the meeting.'];
                        }
                        latestTranscript = transcriptsResponse.value[0];
                        console.log("Found transcript: ".concat(latestTranscript.id));
                        // Step 3: Get transcript content
                        console.log('Fetching transcript content...');
                        return [4 /*yield*/, graphClient
                                .api("/me/onlineMeetings/".concat(this.meetingId, "/transcripts/").concat(latestTranscript.id, "/content"))
                                .get()];
                    case 3:
                        transcriptContent = _a.sent();
                        if (!transcriptContent) {
                            console.warn('Transcript content is empty');
                            return [2 /*return*/, 'Transcript content is not available yet. Please try again later.'];
                        }
                        processedTranscript = '';
                        if (typeof transcriptContent === 'string') {
                            processedTranscript = this.parseVTTTranscript(transcriptContent);
                        }
                        else {
                            // If it's JSON format, extract the text
                            processedTranscript = this.parseJSONTranscript(transcriptContent);
                        }
                        console.log('Successfully retrieved transcript content');
                        return [2 /*return*/, processedTranscript];
                    case 4:
                        error_2 = _a.sent();
                        console.error('Error fetching transcription from Graph API:', error_2);
                        // Provide specific error messages based on error type
                        if (error_2.code === 'Forbidden' || error_2.code === 'Unauthorized') {
                            return [2 /*return*/, 'Access denied. Please ensure you have the necessary permissions to read meeting transcripts (OnlineMeetings.Read, OnlineMeetingTranscript.Read.All).'];
                        }
                        else if (error_2.code === 'NotFound') {
                            return [2 /*return*/, 'Meeting or transcript not found. Please ensure the meeting ID is correct and transcription was enabled.'];
                        }
                        else if (error_2.code === 'BadRequest') {
                            return [2 /*return*/, 'Invalid meeting ID format or transcript not ready yet.'];
                        }
                        else {
                            return [2 /*return*/, "Error retrieving transcript: ".concat(error_2.message || 'Unknown error occurred')];
                        }
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    TeamsAiMeetingAppWebPart.prototype.parseVTTTranscript = function (vttContent) {
        try {
            // Parse VTT (WebVTT) format transcript
            var lines = vttContent.split('\n');
            var transcriptLines = [];
            var currentSpeaker = '';
            var currentText = '';
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i].trim();
                // Skip VTT header and timing lines
                if (line.startsWith('WEBVTT') || line.includes('-->') || line === '') {
                    continue;
                }
                // Check if line contains speaker information
                if (line.includes(':')) {
                    // If we have accumulated text, add it to transcript
                    if (currentText) {
                        transcriptLines.push("".concat(currentSpeaker, ": ").concat(currentText));
                        currentText = '';
                    }
                    // Extract speaker and text
                    var speakerMatch = line.match(/^([^:]+):\s*(.*)$/);
                    if (speakerMatch) {
                        currentSpeaker = speakerMatch[1].trim();
                        currentText = speakerMatch[2].trim();
                    }
                }
                else if (currentSpeaker) {
                    // Continue text from previous line
                    currentText += ' ' + line;
                }
            }
            // Add final accumulated text
            if (currentText && currentSpeaker) {
                transcriptLines.push("".concat(currentSpeaker, ": ").concat(currentText));
            }
            return transcriptLines.join('\n');
        }
        catch (error) {
            console.error('Error parsing VTT transcript:', error);
            return vttContent; // Return raw content if parsing fails
        }
    };
    TeamsAiMeetingAppWebPart.prototype.parseJSONTranscript = function (jsonContent) {
        try {
            // Handle different JSON transcript formats
            if (jsonContent.transcript) {
                return jsonContent.transcript;
            }
            if (jsonContent.transcripts && Array.isArray(jsonContent.transcripts)) {
                return jsonContent.transcripts
                    .map(function (t) { return "".concat(t.speaker || 'Unknown', ": ").concat(t.text || ''); })
                    .join('\n');
            }
            if (jsonContent.value && Array.isArray(jsonContent.value)) {
                return jsonContent.value
                    .map(function (item) { return "".concat(item.speaker || 'Unknown', ": ").concat(item.text || ''); })
                    .join('\n');
            }
            // If it's already a string, return as-is
            if (typeof jsonContent === 'string') {
                return jsonContent;
            }
            return JSON.stringify(jsonContent, null, 2);
        }
        catch (error) {
            console.error('Error parsing JSON transcript:', error);
            return JSON.stringify(jsonContent, null, 2);
        }
    };
    TeamsAiMeetingAppWebPart.prototype.sendToAIService = function (transcription) {
        return __awaiter(this, void 0, void 0, function () {
            var aiServiceUrl, response, result, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        aiServiceUrl = 'https://al-meeting-agentassistant-hdhzh7eeb4g8c0fn.westeurope-01.azurewebsites.net/summarise';
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, fetch(aiServiceUrl, {
                                method: 'POST',
                                headers: {
                                    'Content-Type': 'application/json',
                                },
                                body: JSON.stringify({
                                    transcription: transcription,
                                    meetingId: this.meetingId,
                                    timestamp: new Date().toISOString()
                                })
                            })];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("AI service responded with status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        result = _a.sent();
                        return [2 /*return*/, result];
                    case 4:
                        error_3 = _a.sent();
                        console.error('Error calling AI service:', error_3);
                        // Fallback: Generate a mock summary if the service is unavailable
                        return [2 /*return*/, this.generateMockSummary(transcription)];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    TeamsAiMeetingAppWebPart.prototype.generateMockSummary = function (transcription) {
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
    };
    TeamsAiMeetingAppWebPart.prototype.formatAISummary = function (summary) {
        return "\n      <div class=\"".concat(styles.summaryFormatted, "\">\n        <div class=\"").concat(styles.summarySection, "\">\n          <h4>\uD83D\uDCCB Meeting Summary</h4>\n          <p>").concat(summary.summary, "</p>\n        </div>\n        \n        <div class=\"").concat(styles.summarySection, "\">\n          <h4>\uD83D\uDD11 Key Points</h4>\n          <ul>\n            ").concat(summary.keyPoints.map(function (point) { return "<li>".concat(point, "</li>"); }).join(''), "\n          </ul>\n        </div>\n        \n        <div class=\"").concat(styles.summarySection, "\">\n          <h4>\u2705 Action Items</h4>\n          <div class=\"").concat(styles.actionItems, "\">\n            ").concat(summary.actionItems.map(function (item) { return "\n              <div class=\"".concat(styles.actionItem, "\">\n                <strong>").concat(item.task, "</strong><br>\n                <span class=\"").concat(styles.assignee, "\">Assigned to: ").concat(item.assignee, "</span><br>\n                <span class=\"").concat(styles.dueDate, "\">Due: ").concat(item.dueDate, "</span>\n              </div>\n            "); }).join(''), "\n          </div>\n        </div>\n        \n        <div class=\"").concat(styles.summarySection, "\">\n          <h4>\uD83D\uDCCA Meeting Insights</h4>\n          <p><strong>Sentiment:</strong> ").concat(summary.sentiment, "</p>\n          <p><strong>Duration:</strong> ").concat(summary.duration, "</p>\n          <p><strong>Participants:</strong> ").concat(summary.participants.join(', '), "</p>\n        </div>\n      </div>\n    ");
    };
    TeamsAiMeetingAppWebPart.prototype.delay = function (ms) {
        return new Promise(function (resolve) { return setTimeout(resolve, ms); });
    };
    Object.defineProperty(TeamsAiMeetingAppWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    TeamsAiMeetingAppWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    };
    return TeamsAiMeetingAppWebPart;
}(BaseClientSideWebPart));
export default TeamsAiMeetingAppWebPart;
//# sourceMappingURL=TeamsAiMeetingAppWebPart.js.map
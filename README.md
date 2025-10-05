# Teams AI Meeting App

A SharePoint Framework (SPFx) Teams meeting app that provides AI-powered meeting transcription and summarization capabilities.

## Features

- 📝 **Meeting Information Display**: Shows meeting ID and status when in Teams meetings
- 🎙️ **Transcription Capture**: Retrieves meeting transcriptions post-meeting
- 🤖 **AI Summarization**: Sends transcriptions to custom AI service for intelligent summarization
- 📊 **Rich Summary Display**: Shows key points, action items, and meeting insights in Teams AI assistant style
- 🔗 **Custom API Integration**: Integrates with your AI service at `https://al-meeting-agentassistant-hdhzh7eeb4g8c0fn.westeurope-01.azurewebsites.net/summarise`

## Architecture

- **SPFx Web Part**: Built with SharePoint Framework for seamless Teams integration
- **Teams Context Detection**: Automatically detects meeting context and displays relevant information
- **Microsoft Graph Integration**: Ready for real transcription API integration (currently uses demo data)
- **Custom AI Service**: Sends transcription data to your AI summarization endpoint
- **Responsive UI**: Modern, accessible interface that matches Teams design language

## Prerequisites

- Node.js 16.x (SPFx compatibility requirement)
- SharePoint Framework development environment
- Microsoft Teams admin access for app deployment
- SharePoint tenant with app catalog

## Installation & Setup

### 1. Install Dependencies

```powershell
cd teams-meeting-app
npm install
```

### 2. Build the Solution

```powershell
# Development build
gulp bundle

# Production build
gulp bundle --ship
gulp package-solution --ship
```

### 3. Deploy to SharePoint

1. Upload the generated `.sppkg` file from `sharepoint/solution/` to your SharePoint App Catalog
2. Select "Make this solution available to all sites in the organization"
3. Click "Deploy"

### 4. Sync to Teams

1. In SharePoint App Catalog, select the deployed package
2. Click "Sync to Teams" in the ribbon
3. Confirm the sync operation

### 5. Install in Teams

1. Open Microsoft Teams admin center
2. Go to Teams apps > Manage apps
3. Find "AI Meeting Assistant" and approve it
4. Users can now add the app to their meetings

## Configuration

### Teams App Manifest

The app is configured to work in the following contexts:
- Meeting side panel
- Meeting details tab  
- Meeting chat tab
- Channel tab
- Private chat tab

### API Integration

The app sends POST requests to your AI service with the following payload:

```json
{
  "transcription": "Meeting transcription text...",
  "meetingId": "teams-meeting-id",
  "timestamp": "2025-10-02T17:21:00.000Z"
}
```

Expected response format:
```json
{
  "summary": "Meeting summary text",
  "keyPoints": ["Point 1", "Point 2"],
  "actionItems": [
    {
      "task": "Task description",
      "assignee": "Person name",
      "dueDate": "Due date"
    }
  ],
  "sentiment": "Positive",
  "duration": "30 minutes",
  "participants": ["Name1", "Name2"]
}
```

## Development

### Local Testing

```powershell
# Start development server
gulp serve
```

Navigate to SharePoint workbench to test the web part.

### Teams Testing

1. Create a Teams meeting (not a channel meeting)
2. Invite at least one participant
3. In the meeting details, click the "+" button
4. Select "AI Meeting Assistant" app
5. Test the transcription and summarization features

## Microsoft Graph Integration

For production use, replace the demo transcription with real Microsoft Graph API calls:

```typescript
// Example Graph API call for meeting transcriptions
const graphClient = this.context.msGraphClientFactory.getClient();
const transcripts = await graphClient
  .api(`/me/onlineMeetings/${this.meetingId}/transcripts`)
  .get();
```

Required permissions:
- `OnlineMeetings.Read`
- `OnlineMeetingTranscript.Read.All`

## File Structure

```
teams-meeting-app/
├── src/
│   └── webparts/
│       └── teamsAiMeetingApp/
│           ├── TeamsAiMeetingAppWebPart.ts      # Main web part logic
│           ├── TeamsAiMeetingAppWebPart.module.scss  # Styles
│           ├── TeamsAiMeetingAppWebPart.manifest.json # Web part manifest
│           └── loc/                              # Localization files
├── teams/
│   └── manifest.json                           # Teams app manifest
├── config/                                     # Build configuration
└── package.json                               # Dependencies
```

## Key Components

### TeamsAiMeetingAppWebPart.ts

Main web part class containing:
- Teams context detection
- Meeting transcription handling
- AI service integration
- UI rendering and event handling

### Teams App Manifest

Configured for meeting app contexts with proper scopes and permissions.

### Styling

Modern, responsive design using Office UI Fabric classes and custom SCSS.

## Troubleshooting

### Common Issues

1. **Node.js Version**: Ensure you're using Node.js 16.x for SPFx compatibility
2. **Teams Context**: App requires being added to actual Teams meetings to access meeting context
3. **API CORS**: Ensure your AI service allows cross-origin requests from SharePoint domains
4. **Permissions**: Meeting transcription requires proper Graph API permissions

### Debug Mode

Enable debug mode in serve.json for detailed logging during development.

## Production Considerations

1. **Security**: Implement proper authentication for your AI service
2. **Performance**: Consider caching and rate limiting for API calls
3. **Privacy**: Ensure compliance with data protection regulations
4. **Monitoring**: Add telemetry and error tracking
5. **Graph API**: Replace demo transcription with real Graph API integration

## Support

For issues related to:
- SPFx development: [SharePoint Framework documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/)
- Teams apps: [Microsoft Teams developer documentation](https://docs.microsoft.com/en-us/microsoftteams/platform/)
- Graph API: [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph/)

## License

This project is provided as-is for educational and development purposes.

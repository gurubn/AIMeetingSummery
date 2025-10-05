# Teams AI Meeting App - Deployment Guide

## Overview
This SharePoint Framework (SPFx) Teams meeting app provides:
- Meeting ID display when added to Teams meetings
- Real Microsoft Graph API integration for transcription data
- AI-powered meeting summarization via custom API
- Teams AI summary-style response display

## Prerequisites
- SharePoint Admin Center access
- Teams Admin Center access
- Microsoft Graph API permissions approval rights
- Custom AI API endpoint: `https://al-meeting-agentassistant-hdhzh7eeb4g8c0fn.westeurope-01.azurewebsites.net/summarise`

## Deployment Steps

### Step 1: Upload to SharePoint App Catalog
1. Navigate to SharePoint Admin Center
2. Go to **More features** > **Apps** > **App Catalog**
3. Upload `sharepoint\solution\teams-meeting-app.sppkg`
4. Check **"Make this solution available to all sites in the organization"**
5. Click **Deploy**

### Step 2: Grant Microsoft Graph API Permissions
After deployment, you'll see a notification about API permissions:

1. In SharePoint Admin Center, go to **Advanced** > **API access**
2. Find the pending requests for:
   - **Microsoft Graph** - `OnlineMeetings.Read`
   - **Microsoft Graph** - `OnlineMeetingTranscript.Read.All`
3. **Approve** both permissions

> **Important**: These permissions are required for the app to access real meeting transcription data from Microsoft Graph API.

### Step 3: Sync to Microsoft Teams (Option A - Automatic)
1. In SharePoint Admin Center, go to **More features** > **Apps** > **App Catalog**
2. Find your app and click **"Sync to Teams"**
3. The app will appear in Teams App Store for your organization

### Step 3: Manual Teams App Upload (Option B - Manual)
1. Open Microsoft Teams
2. Go to **Apps** > **Manage your apps** > **Upload an app**
3. Select **Upload an app to your org's app catalog**
4. Upload `teams\TeamsSPFxApp.zip`

## Usage Instructions

### Adding to Teams Meetings
1. In a Teams meeting, click **Apps** (+ button)
2. Search for "Teams AI Meeting App"
3. Click **Add** to add it to the meeting
4. The app will display in the meeting sidebar

### App Functionality
1. **Meeting Detection**: Automatically detects when running in a Teams meeting context
2. **Meeting ID Display**: Shows the current meeting's ID
3. **Transcription Access**: After the meeting ends, fetches real transcription data via Microsoft Graph API
4. **AI Summarization**: Sends transcription to custom AI API for summarization
5. **Results Display**: Shows AI summary in Teams-style formatting

## Technical Implementation

### Microsoft Graph API Integration
- **Authentication**: Uses MSGraphClientV3 from @microsoft/sp-http
- **Endpoints**: 
  - `/me/onlineMeetings` - Get meeting details
  - `/me/onlineMeetings/{meetingId}/transcripts` - Get transcription data
- **Format Support**: VTT and JSON transcript formats
- **Error Handling**: Comprehensive error handling for API failures

### Custom AI API Integration
- **Endpoint**: POST to `/summarise`
- **Payload**: `{ transcript: "meeting transcription text" }`
- **Response**: AI-generated meeting summary
- **Error Handling**: Fallback messages for API failures

### Security Features
- **Scoped Permissions**: Only requests necessary Graph API permissions
- **Error Boundaries**: Graceful degradation when APIs are unavailable
- **User Context**: Respects user's meeting access permissions

## Troubleshooting

### Common Issues

1. **"API permissions not granted" error**
   - Ensure SharePoint Admin has approved Graph API permissions
   - Check API access page in SharePoint Admin Center

2. **"Meeting not found" error**
   - Verify the app is running within a Teams meeting context
   - Check that user has access to the meeting

3. **"Transcription not available" error**
   - Meeting transcription must be enabled for the meeting
   - Transcripts are only available after meeting ends
   - User must have permission to access meeting transcripts

4. **AI summarization fails**
   - Check custom AI API endpoint availability
   - Verify network connectivity to AI service
   - Review browser console for detailed error messages

### Support Information
- **SharePoint Framework Version**: 1.18.2
- **TypeScript Version**: 4.7.4
- **Microsoft Graph Client**: @microsoft/sp-http v1.18.2
- **Teams App Manifest**: v1.13

## File Structure
```
teams-meeting-app/
├── sharepoint/solution/teams-meeting-app.sppkg    # SharePoint package
├── teams/TeamsSPFxApp.zip                         # Teams app package
├── src/webparts/teamsAiMeetingApp/                # Source code
└── config/package-solution.json                   # API permissions config
```

## Next Steps After Deployment
1. Test the app in a Teams meeting
2. Verify Graph API permissions are working
3. Test AI summarization with real meeting data
4. Monitor app performance and user feedback
5. Consider additional features like meeting analytics

---
*Generated: $(Get-Date)*

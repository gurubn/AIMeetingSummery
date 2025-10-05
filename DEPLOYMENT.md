# Deployment Guide for Teams AI Meeting App

## Prerequisites Check

Before deployment, ensure you have:

- [x] SharePoint tenant with App Catalog
- [x] Teams admin access
- [x] Node.js 16.x installed (required for SPFx 1.18.2)
- [x] SharePoint Framework development tools

## Step-by-Step Deployment

### 1. Build the Solution

```powershell
cd teams-meeting-app

# Install dependencies (if not already done)
npm install

# Create production build
gulp bundle --ship
gulp package-solution --ship
```

This creates a `.sppkg` file in the `sharepoint/solution/` directory.

### 2. Create Teams App Package

The Teams app requires a ZIP package containing:
- `manifest.json` (Teams app manifest)
- Icon files (color and outline PNG files)

Since we don't have actual icon files yet, you'll need to:

1. Create two PNG icons:
   - **Color icon**: 96x96 pixels, full color
   - **Outline icon**: 32x32 pixels, white outline on transparent background

2. Name them:
   - `a1b2c3d4-e5f6-789a-bcde-f0123456789a_color.png`
   - `a1b2c3d4-e5f6-789a-bcde-f0123456789a_outline.png`

3. Place icons in the `teams/` folder

4. Create the Teams app package:
   ```powershell
   # Navigate to teams folder
   cd teams
   
   # Create ZIP package (Windows)
   Compress-Archive -Path manifest.json,*.png -DestinationPath ..\src\teams\TeamsSPFxApp.zip
   ```

### 3. Deploy to SharePoint

1. **Access App Catalog**:
   - Go to your SharePoint Admin Center
   - Navigate to "More features" > Apps > App Catalog
   - Or directly visit: `https://[tenant]-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/appCatalog`

2. **Upload SPFx Package**:
   - Upload the `.sppkg` file from `sharepoint/solution/`
   - **Important**: Check "Make this solution available to all sites in the organization"
   - Click "Deploy"

3. **Verify Deployment**:
   - The app should appear in your App Catalog
   - Status should show "Deployed"

### 4. Sync to Teams

1. **In SharePoint App Catalog**:
   - Select your deployed app
   - Click "Sync to Teams" in the ribbon (Files tab)
   - Wait for confirmation message

2. **Alternative Method**:
   - Use the Teams app package ZIP file
   - Upload directly to Teams admin center

### 5. Configure Teams Admin Center

1. **Access Teams Admin Center**:
   - Go to `https://admin.teams.microsoft.com`
   - Navigate to "Teams apps" > "Manage apps"

2. **Find and Approve App**:
   - Search for "AI Meeting Assistant"
   - If status is "Blocked", click to review and approve
   - Set appropriate permission policies

3. **App Policies** (Optional):
   - Navigate to "Teams apps" > "Setup policies"
   - Add the app to relevant policies for automatic installation

### 6. Test Installation

1. **Create Test Meeting**:
   - Create a new Teams meeting (not a channel meeting)
   - Invite at least one participant
   - Save the meeting

2. **Add App to Meeting**:
   - Open meeting details
   - Click the "+" (Add tab) button
   - Search for "AI Meeting Assistant"
   - Add the app

3. **Verify Functionality**:
   - Check that meeting ID is displayed
   - Test the "Get Transcription" button
   - Verify AI summary generation

## Configuration for Production

### Microsoft Graph API Integration

For real transcription access, configure Graph API permissions:

1. **App Registration** (Azure AD):
   ```
   Required Permissions:
   - OnlineMeetings.Read
   - OnlineMeetingTranscript.Read.All
   ```

2. **Update Web Part Code**:
   Replace demo transcription code with real Graph API calls:
   ```typescript
   const graphClient = this.context.msGraphClientFactory.getClient();
   const transcripts = await graphClient
     .api(`/me/onlineMeetings/${this.meetingId}/transcripts`)
     .get();
   ```

### AI Service Configuration

Ensure your AI service at `https://al-meeting-agentassistant-hdhzh7eeb4g8c0fn.westeurope-01.azurewebsites.net/summarise`:

1. **Accepts CORS requests** from SharePoint domains:
   - `*.sharepoint.com`
   - `*.office.com`

2. **Handles the expected payload**:
   ```json
   {
     "transcription": "meeting text...",
     "meetingId": "teams-meeting-id",
     "timestamp": "ISO-date"
   }
   ```

3. **Returns expected response format**:
   ```json
   {
     "summary": "text",
     "keyPoints": ["array"],
     "actionItems": [{"task": "", "assignee": "", "dueDate": ""}],
     "sentiment": "string",
     "duration": "string",
     "participants": ["array"]
   }
   ```

## Troubleshooting

### Common Issues

1. **Node.js Version Error**:
   - SPFx 1.18.2 requires Node.js 18.x, but we've configured for 16.x compatibility
   - If build fails, try using Node.js 18.x

2. **Teams App Not Appearing**:
   - Verify "Sync to Teams" was successful
   - Check Teams admin center for app approval status
   - Wait 10-15 minutes for propagation

3. **Meeting Context Not Available**:
   - App must be installed in actual Teams meetings
   - Channel meetings have different context
   - Test with scheduled meetings, not instant meetings

4. **API CORS Errors**:
   - Configure your AI service to accept requests from SharePoint domains
   - Check browser developer tools for specific CORS errors

5. **Transcription Not Available**:
   - Meeting must be ended for transcription to be available
   - Transcription feature must be enabled in Teams admin settings
   - Currently using demo data - implement real Graph API integration

### Debug Mode

Enable additional logging:

1. Update `serve.json` for local debugging
2. Use browser developer tools to inspect network requests
3. Check SharePoint ULS logs for server-side issues

### Support Resources

- **SPFx Issues**: [SharePoint Framework GitHub](https://github.com/SharePoint/sp-dev-docs)
- **Teams Platform**: [Microsoft Teams Developer Documentation](https://docs.microsoft.com/en-us/microsoftteams/platform/)
- **Graph API**: [Microsoft Graph Documentation](https://docs.microsoft.com/en-us/graph/)

## Security Considerations

1. **Data Privacy**: Ensure compliance with your organization's data policies
2. **API Security**: Implement authentication for your AI service
3. **Permissions**: Use principle of least privilege for Graph API permissions
4. **Audit**: Enable logging for security monitoring

## Next Steps

After successful deployment:

1. **User Training**: Provide guidance on how to use the app
2. **Monitoring**: Set up telemetry and error tracking
3. **Feedback**: Collect user feedback for improvements
4. **Updates**: Plan regular updates and maintenance

## Rollback Plan

If issues occur:

1. **Remove from Teams**: Uninstall app from Teams admin center
2. **Retract SPFx Package**: Remove from SharePoint App Catalog
3. **Fix Issues**: Address problems in development environment
4. **Redeploy**: Follow deployment process again

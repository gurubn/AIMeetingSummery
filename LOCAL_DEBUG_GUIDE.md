# Local Development and Debugging Guide

## 🎯 Current Status

✅ **Development server is running!**  
- **Local server**: https://localhost:4321
- **Workbench URL**: https://localhost:5432/temp/workbench.html?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js
- **LiveReload**: Active on port 35729 (automatic refresh on file changes)

## 🚀 Quick Start - Already Running

The development environment is currently active with:

```powershell
# Already running in terminal
gulp serve --nobuild
```

**Access your web part**:
1. The browser should have opened automatically to the SharePoint workbench
2. If not, navigate to: https://localhost:5432/temp/workbench.html?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js
3. Click the "+" button to add your web part
4. Search for "Teams AI Meeting App" and add it to the page

## 🔧 Development Workflow

### Making Changes

1. **Edit Files**: Make changes to your TypeScript, SCSS, or other source files
2. **Auto-Reload**: The browser will automatically refresh when files change
3. **Debug**: Use browser developer tools for debugging

### Key Files for Development

- **Main Logic**: `src/webparts/teamsAiMeetingApp/TeamsAiMeetingAppWebPart.ts`
- **Styles**: `src/webparts/teamsAiMeetingApp/TeamsAiMeetingAppWebPart.module.scss`
- **Manifest**: `src/webparts/teamsAiMeetingApp/TeamsAiMeetingAppWebPart.manifest.json`

### Testing Features

Since you're not in actual Teams context, the web part will show:
- Welcome panel with feature descriptions
- How-it-works section
- When you deploy to Teams, it will automatically detect meeting context

## 🐛 Debugging Tips

### Browser Developer Tools

1. **Open DevTools**: Press F12 in the browser
2. **Console Tab**: View JavaScript errors and console.log outputs
3. **Network Tab**: Monitor API calls to your AI service
4. **Sources Tab**: Set breakpoints in TypeScript code

### Common Debug Scenarios

#### 1. Testing API Integration
```typescript
// Add debug logging in TeamsAiMeetingAppWebPart.ts
console.log('Sending to AI service:', transcription);
console.log('AI service response:', result);
```

#### 2. Simulating Teams Context
```typescript
// Temporarily modify render() method to test meeting UI
// Change this.isInMeeting to true for testing
this.isInMeeting = true; // Force meeting mode for testing
```

#### 3. Testing Error Handling
```typescript
// Simulate API failures
throw new Error('Test error'); // Add in sendToAIService method
```

### VS Code Debugging

If using VS Code:

1. **Install Extension**: "SharePoint Framework Debug" extension
2. **Set Breakpoints**: Click in the gutter next to line numbers
3. **Launch Debugger**: F5 or Debug menu

## 📱 Testing Teams Context

### Local Testing Limitations

- **No Teams Context**: Local workbench doesn't provide Teams meeting context
- **Mock Data**: Currently uses demo transcription data
- **API Testing**: Can test your AI service integration locally

### Simulating Teams Meeting

To test meeting functionality locally:

1. **Edit TeamsAiMeetingAppWebPart.ts**:
```typescript
// In render() method, add this for testing:
this.meetingId = "test-meeting-123";
this.isInMeeting = true;
```

2. **Save File**: Auto-reload will show meeting interface
3. **Test Features**: Click "Get Post-Meeting Transcription"

### Real Teams Testing

For full Teams context testing:

1. **Deploy to SharePoint**: Follow deployment guide
2. **Add to Teams Meeting**: Install app in actual Teams meeting
3. **Test in Context**: Meeting ID and context will be real

## 🔄 Server Management

### Current Terminal Commands

```powershell
# Server is currently running, but you can also use:

# Stop server: Ctrl+C in the terminal
# Restart server:
gulp serve

# Clean build and serve:
gulp clean
gulp serve

# Build without serving:
gulp build
```

### Common Issues & Solutions

#### Port Already in Use
```powershell
# Kill processes on ports 4321 or 5432
netstat -ano | findstr :4321
taskkill /PID [PID_NUMBER] /F
```

#### Certificate Issues
- Browser may warn about self-signed certificate
- Click "Advanced" → "Proceed to localhost" to continue

#### TypeScript Errors
- Errors appear in terminal output
- Fix errors in source files
- Auto-reload will recompile

## 🧪 Live Testing Features

### Current Working Features

1. **UI Rendering**: ✅ Web part loads and displays correctly
2. **Style Loading**: ✅ SCSS styles are applied
3. **Event Handling**: ✅ Button clicks work
4. **Demo Data**: ✅ Mock transcription and AI response
5. **Error Handling**: ✅ Try/catch blocks and user feedback

### Test Your AI Service

With the server running, you can test your AI endpoint:

1. Click "Get Post-Meeting Transcription" button
2. Watch browser DevTools Network tab
3. POST request goes to: `https://al-meeting-agentassistant-hdhzh7eeb4g8c0fn.westeurope-01.azurewebsites.net/summarise`
4. Response formatted and displayed

### API Request Format Being Sent

```json
{
  "transcription": "John: Good morning everyone...",
  "meetingId": "",
  "timestamp": "2025-10-02T17:59:00.000Z"
}
```

## 📊 Monitoring

### Watch for Changes

The terminal shows:
- **File Changes**: When files are modified
- **Compilation Status**: TypeScript compilation results
- **Errors**: Build or runtime errors
- **Performance**: Build timing information

### Browser Console

Check for:
- **JavaScript Errors**: Red error messages
- **Network Requests**: API calls to your service
- **Custom Logs**: console.log outputs from your code

## 🎉 You're All Set!

Your development environment is ready for:

✅ **Live Development**: Edit files and see changes immediately  
✅ **API Testing**: Test integration with your AI service  
✅ **UI Testing**: Verify styles and interactions  
✅ **Error Testing**: Test error handling scenarios  
✅ **Debug Ready**: Full debugging capabilities available  

**Next Steps:**
1. Open the workbench URL if not already open
2. Add your web part to test the interface
3. Make changes to the source files to see live updates
4. Test the AI service integration
5. When ready, deploy to SharePoint and Teams for full context testing

The server will continue running until you stop it with Ctrl+C in the terminal.

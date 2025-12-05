# Dev Tunnel Setup for Excel Agent

## One-Time Setup (Creates Persistent 30-Day Tunnel)

Follow these steps to create a persistent dev tunnel:

### 1. Login to devtunnel
```powershell
yarn tunnel:login
```
Follow the prompts to authenticate with your Microsoft account.

### 2. Create the tunnel
```powershell
yarn tunnel:create
```
This creates a persistent tunnel that allows anonymous access (required for Microsoft 365 Copilot).

**Save the tunnel ID from the output** - you'll see something like:
```
Tunnel ID: abc123xyz
```

### 3. Add the port
```powershell
yarn tunnel:port
```
This configures the tunnel to forward port 3000 with HTTPS protocol.

### 4. Start the tunnel
```powershell
yarn tunnel:host
```

### 5. Enable the tunnel (First time only)
- Copy the URL labeled "Connect via browser"
- Open it in your browser
- Click "Continue" to enable the tunnel
- **Ignore the error page that appears - this is expected**

### 7. Get your tunnel URL
```powershell
yarn tunnel:show
```
Copy the HTTPS URL (something like: `https://abc123xyz-3000.devtunnels.ms`)

### 7. Update environment variable
Edit `env/.env.dev` and set:
```
DEV_TUNNEL_URL=https://your-tunnel-url-here.devtunnels.ms
```

### 8. Reprovision your app
```powershell
yarn provision
```

## Daily Development Workflow

Once the tunnel is created, you only need to:

1. **Start the tunnel** (in one terminal):
   ```powershell
   yarn tunnel:host
   ```

2. **Start the dev server** (in another terminal):
   ```powershell
   yarn dev-server
   ```

3. **Open Excel** - your agent should now be enabled

## Notes

- The tunnel is **persistent for 30 days** - the URL stays the same
- You can stop the tunnel with `Ctrl+C`
- Restart anytime with `yarn tunnel:host`
- The tunnel URL doesn't change, so you only need to update the manifest once

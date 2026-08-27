import { useEffect } from 'react'
import { PublicClientApplication } from '@azure/msal-browser'

// ---- CONFIG ----
const msalConfig = {
  auth: {
    clientId: '63673e5e-9e45-4bb2-88d5-5a8ed80c4f33', // your Application (client) ID
    authority: 'https://login.microsoftonline.com/b4948e23-f234-4650-9b05-fa1f63e1e6fe', // your Tenant ID
    redirectUri: window.location.origin, // auto-matches whatever URL the app is running on (localhost or Vercel)
  },
}

const sharePointUrl = 'https://kelsons.sharepoint.com/_layouts/15/sharepoint.aspx/discover'

const msalInstance = new PublicClientApplication(msalConfig)
let msalInitialized = false

async function ensureMsalInitialized() {
  if (!msalInitialized) {
    await msalInstance.initialize()
    msalInitialized = true

    // On page load, check if we're returning from a redirect login.
    // If so, this resolves the response and completes the flow.
    const response = await msalInstance.handleRedirectPromise()
    if (response) {
      // Login succeeded — browser now holds a valid Microsoft session cookie,
      // so SharePoint will not prompt again.
      window.location.href = sharePointUrl
    }
  }
}

function App() {
  useEffect(() => {
    // Runs once when the page loads — catches the case where the browser
    // has just come back from Microsoft's login page.
    ensureMsalInitialized()
  }, [])

  const handleSignIn = async () => {
    try {
      await ensureMsalInitialized()
      await msalInstance.loginRedirect({
        scopes: ['User.Read'],
      })
      // Browser navigates away to Microsoft here — code after this line
      // won't run until the user comes back and handleRedirectPromise fires above.
    } catch (err) {
      console.error('Login failed:', err)
      alert('Sign-in failed — please try again.')
    }
  }

  return (
    <div className="page">
      <header>
        <div className="brand">
          <img className="brand-logo" src="/images/logo-m.png" alt="Kelson Innovations" />
        </div>
        <div className="secure">Secure document access</div>
      </header>

      <main>
        <section className="content" aria-labelledby="page-title">
          <p className="eyebrow">Your shared workspace</p>
          <h1 id="page-title">Everything you need, in one place.</h1>
          <p className="intro">You will be required to login to your Microsoft account</p>
        </section>
        <div className="visual" aria-hidden="true">
          <span className="visual-ring visual-ring-one" />
          <span className="visual-ring visual-ring-two" />
          <img className="title-logo" src="/images/logo-title.png" alt="" />
        </div>
        <div className="actions">
          <button className="button" onClick={handleSignIn}>
            Open SharePoint
          </button>
        </div>
      </main>

      <footer>
        <span>Need help accessing your workspace?</span>
        <span>© 2026 Your organization</span>
      </footer>
    </div>
  )
}

export default App
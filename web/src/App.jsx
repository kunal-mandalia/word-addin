import { useEffect, useState } from 'react'
import './App.css'


function App() {
  const [accessToken, setAccessToken] = useState(null);

  useEffect(() => {
    Office.onReady().then(() => {
      console.log("Office is ready");
      
      var options = {
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: false
      };
      
      Office.context.auth.getAccessToken(options, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          var accessToken = asyncResult.value;
          // Use the access token to make authenticated requests
          setAccessToken(accessToken);
        } else {
          console.error(asyncResult.error.message);
        }
      });

    });

    return () => {
      console.log("Cleanup");
    }
  }, []);

  return (
    <>
      <div>
        access token: {accessToken}
      </div>
    </>
  )
}

export default App

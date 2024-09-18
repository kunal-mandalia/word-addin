import { useEffect, useState } from 'react'
import reactLogo from './assets/react.svg'
import viteLogo from '/vite.svg'
import './App.css'


function App() {
  const [count, setCount] = useState(0)

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
          console.log(accessToken);
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
        <a href="https://vitejs.dev" target="_blank">
          <img src={viteLogo} className="logo" alt="Vite logo" />
        </a>
        <a href="https://react.dev" target="_blank">
          <img src={reactLogo} className="logo react" alt="React logo" />
        </a>
      </div>
      <h1>Vite + React</h1>
      <div className="card">
        <button onClick={() => setCount((count) => count + 1)}>
          count is {count}
        </button>
        <p>
          Edit <code>src/App.jsx</code> and save to test HMR
        </p>
      </div>
      <p className="read-the-docs">
        Click on the Vite and React logos to learn more
      </p>
    </>
  )
}

export default App

import { useState, useEffect } from 'react'
import './App.css'

function App() {
  const [officeReady, setOfficeReady] = useState(false);

  useEffect(() => {
    Office.onReady().then(() => {
      console.log("Office is ready");
      setOfficeReady(true);
    });

    return () => {
      console.log("Cleanup");
    }
  }, []);

  return (
    <>
      <div>
        Office is ready: {officeReady ? "Yes" : "No"}
      </div>

      <small>
        v2.0.0
      </small>
    </>
  )
}

export default App

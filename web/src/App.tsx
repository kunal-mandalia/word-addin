import { useState, useEffect } from 'react'
import './App.css'

function App() {
  const [officeReady, setOfficeReady] = useState(false);
  const [storageValue, setStorageValue] = useState<string | null>(null);

  useEffect(() => {
    Office.onReady().then(() => {
      console.log("Office is ready");
      setOfficeReady(true);

      setStorageValue(localStorage.getItem("save_key"));
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

      <div>
        Saved value: {storageValue}{" "}<button onClick={() => {
          const value = new Date().toISOString();
          localStorage.setItem("save_key", value);
          setStorageValue(value);
        }}>Update</button>
      </div>

      <small>
        v2.0.0
      </small>
    </>
  )
}

export default App

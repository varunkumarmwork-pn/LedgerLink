import { useCallback, useEffect, useState } from 'react';
import { Routes, Route } from 'react-router-dom';
import Header from './components/Header.jsx';
import Clients from './pages/Clients.jsx';
import AddClient from './pages/AddClient.jsx';
import Processing from './pages/Processing.jsx';
import Workbook from './pages/Workbook.jsx';
import { deleteClient, listClients } from './api.js';

export default function App() {
  const [clients, setClients] = useState([]);
  const [loadError, setLoadError] = useState(null);
  const [loaded, setLoaded] = useState(false);

  const reload = useCallback(async () => {
    try {
      setClients(await listClients());
      setLoadError(null);
    } catch (e) {
      setLoadError(e.message);
    } finally {
      setLoaded(true);
    }
  }, []);

  useEffect(() => { reload(); }, [reload]);

  const removeClient = async (id) => {
    await deleteClient(id);
    await reload();
  };

  return (
    <div className="app">
      <Header />
      <Routes>
        <Route path="/" element={<Clients clients={clients} loaded={loaded} loadError={loadError} onDelete={removeClient} />} />
        <Route path="/add" element={<AddClient />} />
        <Route path="/processing/:jobId" element={<Processing onFinished={reload} />} />
        <Route path="/client/:id" element={<Workbook clients={clients} onChanged={reload} />} />
      </Routes>
    </div>
  );
}

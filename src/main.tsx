import React, { useCallback } from 'react';

import { Button } from './components/Button';
import execute from './execute';

const { ipcRenderer } = window.require('electron');

const Main: React.FC = () => {
  const openDialog = useCallback(async (): Promise<void> => {
    ipcRenderer.send('send-openDialog');
  }, []);
  return (
    <div>
      <Button onClick={openDialog}>Escolher planilhas</Button>
    </div>

  );
};

export default Main;

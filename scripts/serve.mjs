import { getHttpsServerOptions } from 'office-addin-dev-certs';
import { createPaneServer } from './pane-server.mjs';

createPaneServer(await getHttpsServerOptions()).listen(3000, 'localhost', () => {
  console.log('Formula Trace is serving on https://localhost:3000');
});

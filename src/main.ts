import { register } from './host/plugin-api.js';
import { hasFileApi } from './host/electron.js';
import { VERIFIED_AGAINST } from './host/cardmirror.js';

const ID = 'laymirror';

const ok = register({
  id: ID,
  name: 'laymirror',
  apiVersion: 1,
  commands: [
    {
      id: `${ID}.about`,
      label: 'laymirror: about',
      run: (api) => {
        api.showToast(
          `laymirror — cardmirror ${api.appVersion} ` +
            `(host constants verified against ${VERIFIED_AGAINST}), ` +
            `file api ${hasFileApi() ? 'available' : 'MISSING'}`,
        );
      },
    },
  ],
});

if (!ok) console.warn('[laymirror] __registerCardMirrorPlugin unavailable');

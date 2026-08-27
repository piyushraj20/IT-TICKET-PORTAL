// =====================================================================
// services/pnpjsConfig.ts
// PnPjs is a friendly wrapper over SharePoint's REST API.
// Instead of writing fetch('/_api/web/lists/...') by hand, you write
//   sp.web.lists.getByTitle('ITTickets').items()
//
// PnPjs needs to know WHICH site it is talking to and WHO you are.
// The SPFx "context" object carries both, so we hand it over once,
// at web part startup, and reuse the result everywhere.
// =====================================================================

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, spfi, SPFx } from '@pnp/sp';

// These "side effect" imports switch on parts of the library.
// If you forget one you get errors like "attachmentFiles is not a function".
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';

let _sp: SPFI | undefined;

/**
 * Call once from the web part's onInit with the context,
 * then anywhere else with no arguments.
 */
export const getSP = (context?: WebPartContext): SPFI => {
  if (context) {
    _sp = spfi().using(SPFx(context));
  }
  if (!_sp) {
    throw new Error('PnPjs was used before getSP(context) was called in onInit.');
  }
  return _sp;
};

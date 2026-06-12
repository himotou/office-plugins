import { openLinkBinderDialog } from "../shared/linkBinder";

/* global Office, console */

Office.onReady(() => {
  // Office.js is ready to be called.
});

async function openResourceBinder(event: Office.AddinCommands.Event) {
  try {
    await openLinkBinderDialog();
  } catch (error) {
    console.error("Failed to open resource binder dialog", error);
  } finally {
    event.completed();
  }
}

Office.actions.associate("openResourceBinder", openResourceBinder);

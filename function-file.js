/* global Office */

Office.onReady(() => {
    Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
});

async function onNewMessageCompose(event) {
    try {
        console.log("✅ onNewMessageCompose triggered");

        await Office.addin.showAsTaskpane(); // ✅ This is the correct API to open taskpane
    } catch (error) {
        console.error("❌ Error opening task pane:", error);
    } finally {
        event.completed();
    }
}

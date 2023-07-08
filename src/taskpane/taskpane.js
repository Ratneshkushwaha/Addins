Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
    await Office.addin.showAsTaskpane();
    let data = Office.context.mailbox.item;
    console.log(data.subject);
    // Register an event handler to identify when messages are selected.
    Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, run, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }

      console.log("Event handler added.");
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
          const accessToken = result.value;
          console.log(accessToken);
          // Use the access token.
        } else {
          // Handle the error.
          console.log(accessToken);
        }
      });
    });
  }
});

export async function run() {
  const list = document.getElementById("selected-items");
  while (list.firstChild) {
    console.log(list);
    list.removeChild(list.firstChild);
  }
  const listItem = document.createElement("li");
  let item = Office.context.mailbox.item;
  listItem.textContent = item.subject;
  list.appendChild(listItem);
}

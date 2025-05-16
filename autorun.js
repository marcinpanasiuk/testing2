Office.onReady();

/** inserts a signature automatically when a message is composed or a recipient/sender is changed */
function insertSignature(event) {
    Office.context.mailbox.item.body.setSignatureAsync(`
    <table">
      <tbody>
        <tr>
          <td style="text-align: center">Repro</td>
        </tr>  
      </tbody>
    </table>        
  `, { coercionType: "html" }, function (asyncResult) {
        event.completed();
    });
}

Office.actions?.associate("insertSignature", insertSignature);

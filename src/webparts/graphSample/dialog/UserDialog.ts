import { BaseDialog, IDialogShowOptions } from "@microsoft/sp-dialog";
import styles from "../GraphSampleWebPart.module.scss";
export class UserDialog extends BaseDialog{

    public user: any;

    protected render(){
        
        this.domElement.innerHTML = `<div class='${styles.userDialog}'>
    <h2>User Info</h2>
    <div class="${styles.userDialogInner}">${JSON.stringify(this.user, null, "    ")}</div>
</div>`;
    }

    public static Show(user, config?){
        var dlg = new UserDialog(config);
        dlg.user = user;

        dlg.show();

        return dlg;
    }
}
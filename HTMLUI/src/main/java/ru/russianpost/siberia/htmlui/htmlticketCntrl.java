/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ru.russianpost.siberia.htmlui;

import net.java.html.json.Model;
import net.java.html.json.Function;
import net.java.html.json.Property;
import net.java.html.json.ComputedProperty;
import org.netbeans.api.htmlui.HTMLDialog;
import org.netbeans.api.htmlui.OpenHTMLRegistration;
import org.openide.util.NbBundle;
import org.openide.awt.ActionID;
import org.openide.awt.ActionReference;
import org.openide.awt.ActionReferences;

/**
 * HTML page which displays a window and also a dialog.
 */
@Model(className = "htmlticket", targetId = "", properties = {
    @Property(name = "text", type = String.class)
})
public final class htmlticketCntrl {

    @ComputedProperty
    static String templateName() {
        return "window";
    }

    @Function
    static void showDialog(htmlticket model) {
        String reply = Pages.showhtmlticketDialog(model.getText());
        if ("OK".equals(reply)) {
            model.setText("Happy World!");
        } else {
            model.setText("Sad World!");
        }
    }

    @ActionID(
            category = "Tools",
            id = "ru.russianpost.siberia.htmlui.htmlticket"
    )
    @ActionReferences({
        @ActionReference(path = "Menu/Tools")
        ,
    @ActionReference(path = "Toolbars/File"),})
    @NbBundle.Messages("CTL_htmlticket=Open HTML Hello World!")
    @OpenHTMLRegistration(
            url = "htmlticket.html",
            displayName = "#CTL_htmlticket"
    //, iconBase="SET/PATH/TO/ICON/HERE"
    )
    public static htmlticket onPageLoad() {
        return new htmlticket("Hello World!").applyBindings();
    }

    //
    // dialog UI
    // 
    @HTMLDialog(url = "htmlticket.html")
    static void showhtmlticketDialog(String t) {
        new htmlticketDialog(t, false).applyBindings();
    }

    @Model(className = "htmlticketDialog", targetId = "", properties = {
        @Property(name = "text", type = String.class)
        ,
    @Property(name = "ok", type = boolean.class),})
    static final class DialogCntrl {

        @ComputedProperty
        static String templateName() {
            return "dialog";
        }
    }
}

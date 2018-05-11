/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

import { base64Image } from "./base64Image";

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {

            if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
                console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
            }
            
            $('#insert-image').click(insertImage);
            $('#insert-html').click(insertHTML);

        });
    };


    function insertImage() {
        Word.run(function (context) {
            debugger;
            context.document.body.insertInlinePictureFromBase64(base64Image, "End");

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function insertHTML() {
        Word.run(function (context) {
            
            const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
            blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }  
})();

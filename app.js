/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';
// import { base64Image } from "./base64Image";

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {

            if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
                console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
            }
            
            $('#insert-image').click(insertImage);
        });
    };

		function insertImage() {
			// getBase64ImageFromUrl('https://localhost:3000/assets/icon-32.png')
			getBase64ImageFromUrl('https://localhost:3000/assets/icon-32.png')
		            .then(result => {
		                console.log(result)
		                Word.run(function (context) {
		                    context.document.body.insertInlinePictureFromBase64(result, "End");
		                    return context.sync();
										})
		                .catch(function (error) {
		                    console.log("Error: " + error);
		                    if (error instanceof OfficeExtension.Error) {
		                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
		                    }
										});
								})
								.catch(err => console.error(err));
		}

		// function insertImage() {
		//     Word.run(function (context) {
		        
		//         context.document.body.insertInlinePictureFromBase64(base64Image, "End");

		//         return context.sync();
		//     })
		//     .catch(function (error) {
		//         console.log("Error: " + error);
		//         if (error instanceof OfficeExtension.Error) {
		//             console.log("Debug info: " + JSON.stringify(error.debugInfo));
		//         }
		//     });
		// }

    async function getBase64ImageFromUrl(imageUrl) {
        var res = await fetch(imageUrl);
        var blob = await res.blob();
        console.log(blob)

        return new Promise((resolve, reject) => {
            var reader  = new FileReader();
            reader.addEventListener("load", function () {
                resolve(reader.result);
            }, false);

            reader.onerror = () => {
                return reject(this);
            };
            reader.readAsDataURL(blob);
        })
    }

})();

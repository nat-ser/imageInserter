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
            
            $('.insert-image').click(insertImage);
        });
    };

		function insertImage() {
				//create a canvas element 
				var c = document.createElement("canvas"); 
				var ctx = c.getContext("2d"); 
				var img = document.getElementById("preview"); 
				ctx.drawImage(img, 10, 10); 
				//create the base64 encoded string and take everything after the comma
				//format is normally data:Content-Type;base64,TheStringWeActuallyNeed
				var base64 = c.toDataURL().split(",")[1];
				 
				//insert the TEXT o fthe base64 string into the word document as text
				// Run a batch operation against the Word object model.
				Word.run(function (context) {
				 
				    // Create a proxy object for the document body.
				    var body = context.document.body;
				 
				    // Queue a commmand to insert HTML in to the beginning of the body.
				    body.insertHtml(base64, Word.InsertLocation.start);
				 
				    // Synchronize the document state by executing the queued commands,
				    // and return a promise to indicate task completion.
				    return context.sync().then(function () {
				        console.log('HTML added to the beginning of the document body.');
				    });
				})
				.catch(function (error) {
				    console.log('Error: ' + JSON.stringify(error));
				    if (error instanceof OfficeExtension.Error) {
				        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
				    }
				});
				 
				//Insert the image into the word document as an image created from base64 encoded string
				Word.run(function (context) {
				 
				    // Create a proxy object for the document body.
				    var body = context.document.body;
				 
				    // Queue a command to insert the image.
				    body.insertInlinePictureFromBase64(base64, 'Start');
				 
				    // Synchronize the document state by executing the queued commands,
				    // and return a promise to indicate task completion.
				    return context.sync().then(function () {
				        app.showNotification('Image inserted successfully.');
				    });
				})
				.catch(function (error) {
				    app.showNotification("Error: " + JSON.stringify(error));
				    if (error instanceof OfficeExtension.Error) {
				        app.showNotification("Debug info: " + JSON.stringify(error.debugInfo));
				    }
				});
		}

		// function insertImage() {
		// 	// getBase64ImageFromUrl('https://localhost:3000/assets/icon-32.png')
		// 	getBase64ImageFromUrl('https://localhost:3000/assets/icon-32.png')
		//             .then(result => {
		//                 console.log(result)
		//                 Word.run(function (context) {
		//                     context.document.body.insertInlinePictureFromBase64(result, "End");
		//                     return context.sync();
		// 								})
		//                 .catch(function (error) {
		//                     console.log("Error: " + error);
		//                     if (error instanceof OfficeExtension.Error) {
		//                         console.log("Debug info: " + JSON.stringify(error.debugInfo));
		//                     }
		// 								});
		// 						})
		// 						.catch(err => console.error(err));
		// }



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

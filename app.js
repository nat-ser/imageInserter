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

    // -------------------------------------getBase64FromUrl-------------------------------------
	// function insertImage() {
	// 	getBase64ImageFromUrl('https://localhost:3000/assets/rm.jpg')
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




 //    async function getBase64ImageFromUrl(imageUrl) {
 //        var res = await fetch(imageUrl);
 //        var blob = await res.blob();
 //        console.log(blob)

 //        return new Promise((resolve, reject) => {
 //            var reader  = new FileReader();
 //            reader.addEventListener("load", function () {
 //                resolve(reader.result);
 //            }, false);

 //            reader.onerror = () => {
 //                return reject(this);
 //            };
 //            reader.readAsDataURL(blob);
 //        })
 //    }

    	// ------------------------------------- CANVAS -------------------------------------
		function insertImage() {
				//create a canvas element 
				var canvas = document.createElement("canvas"); 
				var ctx = canvas.getContext("2d"); 
				var img = document.getElementById("preview");
				canvas.width = 400
				canvas.height = 320
				ctx.drawImage(img, 0, 0, (img.width * 4), (img.height * 4)); 


				var base64 = canvas.toDataURL().split(",")[1];
				 
				//Insert the image into the word document as an image created from base64 encoded string
				Word.run(function (context) {
				 
				    // Create a proxy object for the document body.
				    var body = context.document.body;
				 
				    // Queue a command to insert the image.
				    body.insertInlinePictureFromBase64(base64, 'Start');
				 
				    // Synchronize the document state by executing the queued commands,
				    // and return a promise to indicate task completion.
				    return context.sync().then(function () {
				        console.log('Image inserted successfully.');
				    });
				})
				.catch(function (error) {
				    app.showNotification("Error: " + JSON.stringify(error));
				    if (error instanceof OfficeExtension.Error) {
				        console.log("Debug info: " + JSON.stringify(error.debugInfo));
				    }
				});
		}


	// -------------------------------------

	// function insertImage() {
	// 	imgUrl = document.getElementById("preview")
	// 	console.log(imgUrl)
	// 	getBase64ImageFromUrl(imgUrl)
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

 //    async function getBase64ImageFromUrl(imageUrl) {
 //        var res = await fetch(imageUrl);
 //        var blob = await res.blob();
 //        console.log(blob)

 //        return new Promise((resolve, reject) => {
 //            var reader  = new FileReader();
 //            reader.addEventListener("load", function () {
 //                resolve(reader.result);
 //            }, false);

 //            reader.onerror = () => {
 //                return reject(this);
 //            };
 //            reader.readAsDataURL(blob);
 //        })
 //    }



    // -------------------------------------base64 from file -------------------------------------
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
})();

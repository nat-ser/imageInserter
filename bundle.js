/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};

/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {

/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId])
/******/ 			return installedModules[moduleId].exports;

/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			exports: {},
/******/ 			id: moduleId,
/******/ 			loaded: false
/******/ 		};

/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);

/******/ 		// Flag the module as loaded
/******/ 		module.loaded = true;

/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}


/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;

/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;

/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";

/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ function(module, exports) {

	/*
	 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
	 * See LICENSE in the project root for license information.
	 */

	'use strict';
	// import { base64Image } from "./base64Image";

	function _asyncToGenerator(fn) { return function () { var gen = fn.apply(this, arguments); return new Promise(function (resolve, reject) { function step(key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { return Promise.resolve(value).then(function (value) { step("next", value); }, function (err) { step("throw", err); }); } } return step("next"); }); }; }

	(function () {

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

					let getBase64ImageFromUrl = (() => {
									var _ref = _asyncToGenerator(function* (imageUrl) {
													var _this = this;

													var res = yield fetch(imageUrl);
													var blob = yield res.blob();
													console.log(blob);

													return new Promise(function (resolve, reject) {
																	var reader = new FileReader();
																	reader.addEventListener("load", function () {
																					resolve(reader.result);
																	}, false);

																	reader.onerror = function () {
																					return reject(_this);
																	};
																	reader.readAsDataURL(blob);
													});
									});

									return function getBase64ImageFromUrl(_x) {
													return _ref.apply(this, arguments);
									};
					})();

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
									}).catch(function (error) {
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
									}).catch(function (error) {
													app.showNotification("Error: " + JSON.stringify(error));
													if (error instanceof OfficeExtension.Error) {
																	app.showNotification("Debug info: " + JSON.stringify(error.debugInfo));
													}
									});
					}
	})();

/***/ }
/******/ ]);
/* global Word console */

const insertText = async (text: string) => {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
      Promise.resolve();
    });
  } catch (error) {
    console.log("Error: " + error);
    Promise.reject(error);
  }
};

export default insertText;

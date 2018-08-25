'use strict';

const express = require('express');
const docx = require('docx');
const fs = require('fs');

const { DocumentCreator } = require('./models');
const { achievements, education, experiences, skills } = require('./constants');

const app = express();
app.get('/', (req,  res) => {
    const creator = new DocumentCreator();
    const doc = creator.create([experiences, education, skills, achievements]);

    const packer = new docx.Packer();

    packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("My Document.docx", buffer);
    });
});

app.listen(5555);
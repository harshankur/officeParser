const { parseWord, parsePowerPoint, parseExcel, parseOpenOffice, parseOffice } = require('./officeParser');

var parseWordAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseWord(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
    })
}

var parsePowerPointAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parsePowerPoint(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
    })
}

var parseExcelAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseExcel(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
    })
}

var parseExcelAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseExcel(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
    })
}

var parseExcelAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseExcel(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
    })
}

var parseOpenOfficeAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseOpenOffice(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
    })
}

var parseOfficeAsync = function (filename, deleteOfficeDist = true) {
    return new Promise((resolve, reject) => {
        try {
            parseOffice(filename, function (data, err) {
                if (err) return reject(err);
                return resolve(data);
            },deleteOfficeDist);
        } catch (error) {
            return reject(error);
        }
    })
}

module.exports.parseWordAsync = parseWordAsync;
module.exports.parsePowerPointAsync = parsePowerPointAsync;
module.exports.parseExcelAsync = parseExcelAsync;
module.exports.parseOpenOfficeAsync = parseOpenOfficeAsync;
module.exports.parseOfficeAsync = parseOfficeAsync;

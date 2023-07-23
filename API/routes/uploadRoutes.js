const express = require("express")
const multer = require("multer")
const path = require("path")
const { upload_excel } = require("../controllers/uploadController")

const router = express.Router()

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'sheets/uploads')
    },
    filename: function (req, file, cb) {
        cb(null, path.basename(file.originalname))
    }
})
// const storage = multer.memoryStorage()
const upload = multer({ storage: storage })

router.route('/excel').post(upload.single('upload_file'), upload_excel)

module.exports = router
const express = require("express")
const dotenv = require("dotenv")
const uploadRoutes = require("./routes/uploadRoutes")
const reportRoutes = require("./routes/reportRoutes")

const app = express()

app.use(express.json())
app.use(express.urlencoded({extended: true}))
dotenv.config()

app.use((req, res, next) => {
    res.header("Access-Control-Allow-Origin", "*")
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept")
    next();
  })

app.use((req, res, next) => {
    console.log(`${req.method} ${req.url}`)
    next()
})

app.get('/', (req, res) => {
    res.send("Hello")
})

app.use('/upload', uploadRoutes)
app.use('/get-report', reportRoutes)

app.listen(process.env.PORT, () => {
    console.log(`Listening on port ${process.env.PORT}`)
})
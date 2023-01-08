var express    = require('express');
var mongoose   = require('mongoose');
var bodyParser = require('body-parser');
var path       = require('path');

//PACKAGE FOR EXTRACTING DATA FROM EXCEL FILES
var XLSX       = require('xlsx');
var multer     = require('multer');

const app = express();

app.set('view engine', 'ejs');
app.use(express.static(path.resolve(__dirname,'public')));
app.use(bodyParser.urlencoded({limit: '50mb', extended:true}));

const PORT = process.env.PORT || 3000;
const DB = 'mongodb://localhost:27017/excel';

mongoose.connect(DB)
        .then((result) => app.listen(PORT, () => console.log(`Running on port ${PORT}`)))
        .catch((err) => console.log(err));

//MULTER
var storage = multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, './public/uploads')
    },
    filename: function (req, file, cb) {
      cb(null, file.originalname)
    }
  });
  
var upload = multer({ storage: storage });

//SCHEMA AND MODEL

var excelSchema = new mongoose.Schema({
    BodyNumber: {
        type: String
    },
    Name: {
        type: String
    },
    TODA: {
        type: String
    },
    ContactNumber: {
        type: String
    }
});

var excelModel = mongoose.model('driver',excelSchema);

app.get('/', (req, res) => {
    // res.render('index');

    excelModel.find((err,data)=>{
        if(err){
            console.log(err)
        }else{
            if(data!=''){
                res.render('index',{result:data});
            }else{
                res.render('index',{result:{}});
            }
        }
    });
});

app.post('/',upload.single('excel'),(req,res)=>{

    //READS THE FILE SUBMITTED BY THE USER  
    var workbook =  XLSX.readFile(req.file.path);
    var sheet_namelist = workbook.SheetNames;

    var x=0;

    //TO LOOP AROUND EACH SHEETS OF THE UPLOADED EXCEL FILE
    sheet_namelist.forEach(element => {
        
        //CONVERTS EXCEL FILE TO JSON
        //RETURNS JSON
        var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_namelist[x]]);

        //LOOPS AROUND THE JSON TO INSERT EACH OBJECT TO THE DATABASE
        for(let driver in xlData) {
            excelModel.insertMany(xlData[driver]);
        }
        x++;
    });

    res.redirect('/');
  });

const { Sequelize } = require('sequelize');

const database = new Sequelize('database', 'username', 'password' , 'dialect', {
    
    dialect: 'postgres',

    host: '143.198.237.19',

    port: 5432,
  
    username : 'postgres',

    password : 'SArSG9SNBrEJh5LXtNHAVD3GSMvLAkgsVAjpS8Q2',

    database : 'TEST'
  })
  
  export default database
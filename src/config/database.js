import { config } from "dotenv";
import oracledb from "oracledb";
config();

export class DatabaseService {
  constructor() {}

  static getDbConfig() {
    let user;
    let password;
    let connString;
    if (process.env.ENVIROMENT === "INTRA") {
      user = process.env.INTRA_DATABASE_USER;
      password = process.env.INTRA_DATABASE_PASSWORD;
      connString = process.env.INTRA_DATABASE_CONN_STRING;
    } else {
      throw new Error("something went wrong check the ENVIROMENT setup");
    }

    return { user, password, connString };
  }
  static createPool() {
    const pool = oracledb.createPool({
      user: user,
      password: password,
      connectionString: connString,
      poolMin: 5,
      poolMax: 5,
      poolIncrement: 5,
    });
    return pool;
  }
}

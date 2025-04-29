const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const mysql = require('mysql2/promise');
const app = express();
const port = 3000;

// 配置 MySQL 连接
const pool = mysql.createPool({
    host: 'localhost',
    user: 'your_username',
    password: 'your_password',
    database: 'your_database',
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0
});

// 配置文件上传
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// 处理文件上传
app.post('/upload', upload.single('file'), async (req, res) => {
    try {
        const buffer = req.file.buffer;
        const workbook = XLSX.read(buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);

        // 插入或更新数据到数据库
        const connection = await pool.getConnection();
        for (const row of data) {
            const [result] = await connection.execute(
                'INSERT INTO grades (学年, 考试类型, 年级, 学号, 姓名, 班级, 语文, 数学, 英语, 物理, 化学, 道德与法治, 历史, 地理, 生物, 体育, 音乐, 美术, 信息) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON DUPLICATE KEY UPDATE 学年 = VALUES(学年), 考试类型 = VALUES(考试类型), 年级 = VALUES(年级), 姓名 = VALUES(姓名), 班级 = VALUES(班级), 语文 = VALUES(语文), 数学 = VALUES(数学), 英语 = VALUES(英语), 物理 = VALUES(物理), 化学 = VALUES(化学), 道德与法治 = VALUES(道德与法治), 历史 = VALUES(历史), 地理 = VALUES(地理), 生物 = VALUES(生物), 体育 = VALUES(体育), 音乐 = VALUES(音乐), 美术 = VALUES(美术), 信息 = VALUES(信息)',
                [
                    row['学年'],
                    row['考试类型'],
                    row['年级'],
                    row['学号'],
                    row['姓名'],
                    row['班级'],
                    row['语文'],
                    row['数学'],
                    row['英语'],
                    row['物理'],
                    row['化学'],
                    row['道德与法治'],
                    row['历史'],
                    row['地理'],
                    row['生物'],
                    row['体育'],
                    row['音乐'],
                    row['美术'],
                    row['信息']
                ]
            );
        }
        connection.release();

        res.json({ success: true, message: '数据上传成功' });
    } catch (error) {
        console.error('文件处理出错:', error);
        res.status(500).json({ success: false, message: '服务器错误' });
    }
});

// 处理查询请求
app.get('/search', async (req, res) => {
    try {
        const { searchId, searchName, startYear, endYear, searchType } = req.query;
        let query = 'SELECT * FROM grades WHERE 1=1';
        const values = [];

        if (searchId) {
            query += ' AND 学号 LIKE ?';
            values.push(`%${searchId}%`);
        }
        if (searchName) {
            query += ' AND 姓名 LIKE ?';
            values.push(`%${searchName}%`);
        }
        if (startYear) {
            query += ' AND 学年 >= ?';
            values.push(startYear);
        }
        if (endYear) {
            query += ' AND 学年 <= ?';
            values.push(endYear);
        }
        if (searchType) {
            query += ' AND 考试类型 = ?';
            values.push(searchType);
        }

        const connection = await pool.getConnection();
        const [results] = await connection.execute(query, values);
        connection.release();

        res.json({ success: true, results: results });
    } catch (error) {
        console.error('查询出错:', error);
        res.status(500).json({ success: false, message: '服务器错误' });
    }
});

// 处理成绩更新请求
app.post('/update', express.json(), async (req, res) => {
    try {
        const { id, column, newValue } = req.body;
        const query = `UPDATE grades SET ${column} = ? WHERE id = ?`;
        const connection = await pool.getConnection();
        const [result] = await connection.execute(query, [newValue, id]);
        connection.release();

        if (result.affectedRows > 0) {
            res.json({ success: true, message: '成绩更新成功' });
        } else {
            res.json({ success: false, message: '未找到要更新的记录' });
        }
    } catch (error) {
        console.error('更新出错:', error);
        res.status(500).json({ success: false, message: '服务器错误' });
    }
});

app.listen(port, () => {
    console.log(`服务器运行在 http://localhost:${port}`);
});
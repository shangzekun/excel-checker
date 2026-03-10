# Excel Checker（加固版）

这版已经补齐了除真实规则外的主要加固项：

- Excel 文件内容合法性校验
- 非法/损坏文件统一返回 400
- 规则执行异常兜底
- 规则 ID 合法性过滤
- 基础日志
- 环境变量配置
- CORS 可配置
- Docker 非 root 用户运行
- 基础接口测试

## 主要仍需你补充的内容
- `app/checks/` 目录中的真实业务规则

## 本地运行
```bash
pip install -r requirements.txt
uvicorn app.main:app --reload
```

打开：
`http://127.0.0.1:8000`

## 运行测试
```bash
pytest
```

## Docker
```bash
docker compose up --build
```

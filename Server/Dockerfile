# 使用 Miniconda 作为基础镜像
FROM continuumio/miniconda3

# 设置工作目录
WORKDIR /usr/src/app

# 复制环境文件和程序文件到容器中
COPY . .

# 创建 Conda 环境
RUN conda env create -f environment.yml

# 激活环境，并将其设置为默认环境
SHELL ["conda", "run", "-n", "FunASR_Py", "/bin/bash", "-c"]

# 暴露服务端口
EXPOSE 5000

# 运行 Python 程序
CMD ["conda", "run", "--no-capture-output", "-n", "FunASR_Py", "python", "server.py"]

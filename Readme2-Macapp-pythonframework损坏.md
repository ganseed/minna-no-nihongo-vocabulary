# # “Python.framework”已损坏
<p>
  mac客户端运行时出现，“Python.framework”已损坏，无法打开。 你应该将它移到废纸篓。

</p>

---

- 打开 Mac 的 “终端”（Terminal）应用。

- 输入命令 sudo xattr -r -d com.apple.quarantine （注意最后有一个空格）。

- 将报错的软件或包含该 Python.framework 的文件夹直接拖入终端窗口中，回车运行。

- 输入你的 Mac 开机密码（输入时密码不会显示，输完直接按回车）。
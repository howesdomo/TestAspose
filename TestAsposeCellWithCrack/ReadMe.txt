破解版要在Web程序中正常使用, 加上配置

<configuration>
    <system.web>
	.....
      <hostingEnvironment shadowCopyBinAssemblies="false" />
	....
    </system.web>
</configuration>
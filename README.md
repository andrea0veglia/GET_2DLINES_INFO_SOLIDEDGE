# GET 2DLINES INFO IN SOLIDEDGE

The program creates a text file with the information of the lines of the 2D drawing, opened in SolidEdge.

![](https://raw.githubusercontent.com/andrea0veglia/GET_2DLINES_INFO_SOLIDEDGE/master/SE.png)

### - Install Interop.SolidEdge

In order to implement the program in your solution, you need to download Interop.SolidEdge assembly. 

The Interop.SolidEdge assembly is published as a [NuGet](https://www.nuget.org/) package. The package id is [Interop.SolidEdge](https://www.nuget.org/packages/Interop.SolidEdge). The package includes .NET 2.0 & 4.0 builds of the assembly. Depending on your project settings, NuGet will reference the appropriate assembly.

The [Nuget Package Manager Console](http://docs.nuget.org/docs/start-here/using-the-package-manager-console) provides a command line style interface to interact with NuGet. Note that the steps will vary depending on your version of Visual Studio.

![](https://raw.githubusercontent.com/SolidEdgeCommunity/Interop.SolidEdge/master/media/Install.png)


### - Program steps

For the correct functioning of the program it is necessary to follow the following steps:
1) Open a 2D drawing in SolidEdge;
2) Keep the window with the selected design active;
3) Start the program;
4) The log file, which contains the information, will be saved in the "LOGS" folder;


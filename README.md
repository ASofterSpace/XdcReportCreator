# XDC Report Creator

**Class:** Right Internal

**Language:** PHP + Java

**Platform:** Web

This is a small web application that enables the creation of an XDC report based on some user input, with the web part being written in PHP, but the actual report generation being written in Java (which is supposed to then run on the server)

## Setup

1. This requires PHP.
To just locally install it under Windows you can go to [php.net](https://windows.php.net/download) and download the latest zip file. After you unzipped it locally, e.g. to `C:\php\`, you can call in a terminal:
```
C:\php\php -v
```
to ensure that you successfully got it.

2. Now configure the XDC Report Creator by creating a file inside this repository called `php-path.config` which contains just the local PHP path - in this case, `C:\php` - and nothing else.

3. For the Java part, you need to have a JDK of Java 7 or higher running and `java` as well as `javac` on the PATH.

4. Download our Toolbox-Java (which is a separate project here on github) into an adjacent directory on your hard drive.

5. Start the build of the Java part by calling under Windows:

```
build.bat
```

Or under Linux:

```
./build.sh
```

6. Finally, this requires an XLSX template file that will be edited in the process. Put your template file as `template.xlsx` into this repository.

## Run

To start up the XDC Report Creator and let it generate a single report, you can call under Windows:

```
test.bat
```

## License

We at A Softer Space really love the Unlicense, which pretty much allows anyone to do anything with this source code.
For more info, see the file UNLICENSE.

If you desperately need to use this source code under a different license, [contact us](mailto:moya@asofterspace.com) - I am sure we can figure something out.
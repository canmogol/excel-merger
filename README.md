# Excel File Merger

A Java Native command line tool to merge excel files.

## Compile

Here are the steps to native compile.

**How to Compile Native Executable**

You need GraalVM to compile to native application. One way to install GraalVM is to use SDKMan.
You can install the SDK Man from its site https://sdkman.io/install
After that, you can install GraalVM with this command.

```bash
sdk install java 21.3.0.r17-grl
```
You can compile the native executable by running the following command.
```bash
# this should create an executable called 'list' under the 'target' folder
mvn -DbuildArgs=--no-server clean verify
```
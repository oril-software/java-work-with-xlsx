## Processing .xlsx files with Java
Code sample in this repository demonstrates how to create and parse .xlsx files using Java.

**Required Dependencies**

```
<dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi</artifactId>
      <version>5.2.3</version>
</dependency>
<dependency>
      <groupId>org.apache.poi</groupId>
      <artifactId>poi-ooxml</artifactId>
      <version>5.2.3</version>
</dependency>
```

### How to run

Open `XlsxProcessorTest.class` and run `testXlsxProcessing()` method.
It will create and populate `users.xlsx` file and then will parse the same file.

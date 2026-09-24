# Build, package, and install for another project

This guide builds the library from source and installs it into the developer's own local Maven repository. The consumer can then use it as a normal Maven dependency without copying JAR files by hand.

## 1. Build and verify this library

Install JDK 17+ and Maven 3.8+. In a terminal, change to the directory containing this project's `pom.xml`, then run:

```shell
mvn clean verify
```

This compiles the source and runs unit tests. To create the JAR explicitly:

```shell
mvn package
```

The artifact is created at `target/common-csv-library-1.0.0-SNAPSHOT.jar`.

## 2. Install into your local Maven repository

Run this in the library project directory:

```shell
mvn clean install
```

Maven installs the JAR and its POM into the local repository for the current operating-system user (normally `~/.m2/repository`). The library does not include its dependencies inside a shaded JAR; Maven resolves Apache Commons CSV and Apache POI from the POM when the consuming project builds.

## 3. Add it to your own Maven project

In the consuming project's `pom.xml`, add the dependency inside `<dependencies>`:

```xml
<dependency>
  <groupId>io.github.commoncsv</groupId>
  <artifactId>common-csv-library</artifactId>
  <version>1.0.0-SNAPSHOT</version>
</dependency>
```

Then refresh Maven in your IDE or build that project:

```shell
mvn clean compile
```

The project can now import `io.github.commoncsv.CsvInvoiceCalculator` and other public classes. See the examples in [README.md](README.md).

## 4. Install a changed version

When source changes, increment `<version>` in this library's `pom.xml`, then run `mvn clean install` again. Update the dependency version in the consuming project to match. Maven caches released versions, so use a new version when publishing changes rather than replacing an already released version.

## Sharing with a team

Local installation only makes the artifact available to the current developer account and machine. For a team, publish a reviewed version to an organization-approved Maven repository and configure consumers to use it. Set the organization's own `groupId`, choose a release version, and manage repository credentials and signing through the organization's build and release process. This repository intentionally does not publish anything or include personal Maven credentials.

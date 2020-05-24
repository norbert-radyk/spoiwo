lazy val pomDetails = <url>https://github.com/norbert-radyk/spoiwo/</url>
  <licenses>
    <license>
      <name>MIT-License</name>
      <url>http://www.opensource.org/licenses/mit-license.php</url>
      <distribution>repo</distribution>
    </license>
  </licenses>
  <scm>
    <url>https://github.com/norbert-radyk/spoiwo</url>
    <developerConnection>scm:git:git@github.com:norbert-radyk/spoiwo.git</developerConnection>
    <connection>scm:git:git@github.com:norbert-radyk/spoiwo.git</connection>
  </scm>
  <developers>
    <developer>
      <id>norbert-radyk</id>
      <name>Norbert Radyk</name>
      <email>norbert.radyk@gmail.com</email>
    </developer>
  </developers>

lazy val commonSettings = Seq(
  organization := "com.norbitltd",
  scalaVersion := "2.13.2",
  scalacOptions ++= Seq(
    "-deprecation",
    "-feature",
    "-Ywarn-dead-code",
    "-Ywarn-unused",
    ),
  crossScalaVersions := Seq(scalaVersion.value, "2.12.10", "2.11.12"),
  publishMavenStyle := true,
  publishArtifact in Test := false,
  resolvers += "Apache Releases" at "https://repository.apache.org/content/repositories/releases",
  version := "1.8.1-SNAPSHOT",
  publishTo := sonatypePublishToBundle.value,
  pomExtra := pomDetails,
  libraryDependencies ++= Seq(
    "org.scala-lang.modules" %% "scala-xml"         % "1.3.0",
    "org.apache.poi"         %  "poi"               % "4.1.2",
    "org.apache.poi"         %  "poi-ooxml"         % "4.1.2",
    "org.scalatest"          %% "scalatest"         % "3.1.2"   % Test
  )
)

lazy val root = project
  .in(file("."))
  .settings(commonSettings : _*)
  .settings(
    publishArtifact := false
  )
  .aggregate(spoiwo, spoiwoGrids, examples)

lazy val spoiwo = (project in file("core"))
  .settings(commonSettings : _*)
  .settings(
    name := "spoiwo"
  )

lazy val examples = (project in file("examples"))
  .dependsOn(spoiwo, spoiwoGrids)
  .settings(commonSettings : _*)
  .settings(
    name := "spoiwo-examples"
  )

lazy val spoiwoGrids = (project in file("grids"))
  .dependsOn(spoiwo)
  .settings(commonSettings : _*)
  .settings(
    name := "spoiwo-grids"
  )

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
  scalaVersion := "2.12.7",
  scalacOptions ++= Seq(
    "-deprecation",
    "-feature",
    "-Ywarn-dead-code",
    "-Ywarn-unused",
    //"-Ywarn-unused-import"),
    ),
  crossScalaVersions := Seq("2.13.0-RC1", scalaVersion.value, "2.11.12"),
  publishMavenStyle := true,
  publishArtifact in Test := false,
  useGpg := true,
  resolvers += "Apache Releases" at "https://repository.apache.org/content/repositories/releases",
  publishTo := {
    val nexus = "https://oss.sonatype.org/"
    if (isSnapshot.value)
      Some("snapshots" at nexus + "content/repositories/snapshots")
    else
      Some("releases"  at nexus + "service/local/staging/deploy/maven2")
  },
  pomExtra := pomDetails,
  libraryDependencies ++= Seq(
    "org.scala-lang.modules" %% "scala-xml"         % "1.2.0",
    "joda-time"              %  "joda-time"         % "2.9.9",
    "org.joda"               %  "joda-convert"      % "2.0.1",
    "org.apache.poi"         %  "poi"               % "4.1.0",
    "org.apache.poi"         %  "poi-ooxml"         % "4.1.0",
    "org.scalatest"          %% "scalatest"         % "3.0.8-RC2"   % Test
  )
)

lazy val spoiwo = (project in file("."))
  .settings(commonSettings : _*)
  .settings(
    name := "spoiwo",
    version := "1.4.1"
  )

lazy val examples = (project in file("examples"))
  .dependsOn(spoiwo)
  .settings(commonSettings : _*)
  .settings(
    name := "spoiwo-examples",
    version := "1.4.1"
  )

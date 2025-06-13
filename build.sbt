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
    </developer>
  </developers>

lazy val commonSettings = Seq(
  organization := "com.norbitltd",
  ThisBuild / crossScalaVersions := Seq("2.13.16", "3.6.3"),
  ThisBuild / scalaVersion := crossScalaVersions.value.last,
  ThisBuild / githubWorkflowPublishTargetBranches := List(),
  ThisBuild / githubWorkflowBuildPreamble ++= List(),
  scalacOptions ++= Seq(
    "-deprecation",
    "-feature"
  ),
  publishMavenStyle := true,
  Test / publishArtifact := false,
  resolvers += "Apache Releases" at "https://repository.apache.org/content/repositories/releases",
  version := "2.3.0-SNAPSHOT",
  versionScheme := Some("early-semver"),
  publishTo := sonatypePublishToBundle.value,
  pomExtra := pomDetails,
  libraryDependencies ++= Seq(
    "com.github.tototoshi" %% "scala-csv" % "2.0.0",
    "org.apache.poi" % "poi" % "5.4.1",
    "org.apache.poi" % "poi-ooxml" % "5.4.1",
    "org.scalatest" %% "scalatest" % "3.2.19" % Test
  )
)

lazy val root = project
  .in(file("."))
  .settings(commonSettings *)
  .settings(
    publishArtifact := false
  )
  .aggregate(spoiwo, spoiwoGrids, examples)

lazy val spoiwo = (project in file("core"))
  .settings(commonSettings *)
  .settings(
    name := "spoiwo"
  )

lazy val examples = (project in file("examples"))
  .dependsOn(spoiwo, spoiwoGrids)
  .settings(commonSettings *)
  .settings(
    name := "spoiwo-examples"
  )

lazy val spoiwoGrids = (project in file("grids"))
  .dependsOn(spoiwo)
  .settings(commonSettings *)
  .settings(
    name := "spoiwo-grids"
  )

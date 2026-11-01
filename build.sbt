inThisBuild(
  List(
    homepage := Some(url("https://github.com/norbert-radyk/spoiwo")),
    licenses := List("MIT" -> url("http://www.opensource.org/licenses/mit-license.php")),
    scmInfo := Some(
      ScmInfo(
        url("https://github.com/norbert-radyk/spoiwo"),
        "scm:git:git@github.com:norbert-radyk/spoiwo.git",
        "scm:git:git@github.com:norbert-radyk/spoiwo.git"
      )
    ),
    developers := List(Developer("norbert-radyk", "Norbert Radyk", "", url("https://github.com/norbert-radyk")))
  )
)

lazy val commonSettings = Seq(
  ThisBuild / crossScalaVersions := Seq("2.13.18", "3.10.0"),
  ThisBuild / scalaVersion := crossScalaVersions.value.last,
  ThisBuild / githubWorkflowTargetTags ++= Seq("v*"),
  ThisBuild / githubWorkflowPublishTargetBranches := Seq(RefPredicate.StartsWith(Ref.Tag("v"))),
  ThisBuild / githubWorkflowPublish := Seq(
    WorkflowStep.Sbt(
      commands = List("ci-release"),
      env = Map(
        "PGP_PASSPHRASE" -> "${{ secrets.PGP_PASSPHRASE }}",
        "PGP_SECRET" -> "${{ secrets.PGP_SECRET }}",
        "SONATYPE_PASSWORD" -> "${{ secrets.SONATYPE_PASSWORD }}",
        "SONATYPE_USERNAME" -> "${{ secrets.SONATYPE_USERNAME }}"
      )
    )
  ),
  ThisBuild / githubWorkflowBuildPreamble ++= List(),
  ThisBuild / githubWorkflowJavaVersions := Seq(JavaSpec.temurin("17")),
  scalacOptions ++= Seq(
    "-deprecation",
    "-feature"
  ),
  Test / publishArtifact := false,
  resolvers += "Apache Releases" at "https://repository.apache.org/content/repositories/releases",
  versionScheme := Some("early-semver"),
  libraryDependencies ++= Seq(
    "com.github.tototoshi" %% "scala-csv" % "2.0.0",
    "org.apache.poi" % "poi" % "6.0.0",
    "org.apache.poi" % "poi-ooxml" % "6.0.0",
    "org.scalatest" %% "scalatest" % "3.2.20" % Test
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

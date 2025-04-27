plugins {
    kotlin("jvm") version "1.9.23"
}

group = "ru.jengle88"
version = "1.0"

repositories {
    mavenCentral()
}

dependencies {
    implementation ("org.apache.poi:poi:4.1.0")
    implementation ("org.apache.poi:poi-ooxml:4.1.0")
    testImplementation("org.jetbrains.kotlin:kotlin-test")

}

tasks.jar.configure {
    manifest {
        attributes(mapOf("Main-Class" to "ru.jengle88.MainKt"))
    }
    configurations["compileClasspath"].forEach { file: File ->
        from(zipTree(file.absoluteFile))
    }
    duplicatesStrategy = DuplicatesStrategy.INCLUDE
}

tasks.test {
    useJUnitPlatform()
}
kotlin {
    jvmToolchain(17)
}
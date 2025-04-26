plugins {
    kotlin("jvm") version "1.9.23"
}

group = "ru.jengle88"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    implementation ("org.apache.poi:poi:4.1.0")
    implementation ("org.apache.poi:poi-ooxml:4.1.0")
    testImplementation("org.jetbrains.kotlin:kotlin-test")

}

tasks.test {
    useJUnitPlatform()
}
kotlin {
    jvmToolchain(17)
}
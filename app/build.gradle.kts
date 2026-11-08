plugins {
    id("com.android.application")
}

android {
    namespace = "com.wordtopdf.app"
    compileSdk = 35

    defaultConfig {
        applicationId = "com.wordtopdf.app"
        minSdk = 24
        targetSdk = 35
        versionCode = 1
        versionName = "1.0"
    }

    buildTypes {
        release {
            isMinifyEnabled = false
            // Signed with the debug key so the release APK is installable for testing.
            // Replace with your own keystore before publishing to Google Play.
            signingConfig = signingConfigs.getByName("debug")
        }
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }
}

// No third-party dependencies: the .docx reader and PDF writer use only Android/Java APIs.

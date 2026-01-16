package com.sca.macros;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

class TranslationSmokeTest {
    @Test
    void pythonTranslationFilesExist() {
        Path root = Path.of("original-report-form-macros");
        assertTrue(Files.exists(root.resolve("Module2.py")), "Module2.py should exist");
        assertTrue(Files.exists(root.resolve("Module5.py")), "Module5.py should exist");
        assertTrue(Files.exists(root.resolve("Module7.py")), "Module7.py should exist");
    }
}

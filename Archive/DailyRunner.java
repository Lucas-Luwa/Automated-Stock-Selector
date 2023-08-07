import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

public class DailyRunner {
    public static void main(String[] args) {
        try {            
            String method1Name = "goofy";
            String method2Name = "lessgoofy";
            // Create and start the threads
            Thread thread1 = new PyThread("Thread 1", "DailyUpdates.py", method1Name);
            Thread thread2 = new PyThread("Thread 2", "DailyUpdates.py", method2Name);
            Thread thread1 = new PyThread("Thread 1", "DailyUpdates.py", method1Name);
            Thread thread2 = new PyThread("Thread 2", "DailyUpdates.py", method2Name);

            thread1.start();
            thread2.start();
            
            // Wait for the threads to complete
            thread1.join();
            thread2.join();
            
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }
}
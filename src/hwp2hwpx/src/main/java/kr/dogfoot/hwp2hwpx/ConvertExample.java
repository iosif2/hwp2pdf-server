package kr.dogfoot.hwp2hwpx;

import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.reader.HWPReader;
import kr.dogfoot.hwpxlib.object.HWPXFile;
import kr.dogfoot.hwpxlib.writer.HWPXWriter;

/**
 * HWP 파일을 HWPX 파일로 변환하는 예제 클래스
 * 
 * 사용법:
 *   java kr.dogfoot.hwp2hwpx.ConvertExample <입력파일경로> <출력파일경로>
 * 
 * 예제:
 *   java kr.dogfoot.hwp2hwpx.ConvertExample test/빈파일/from.hwp test/빈파일/to.hwpx
 */
public class ConvertExample {
    public static void main(String[] args) {
        try {
            // 명령줄 인자 확인
            if (args.length < 2) {
                System.out.println("사용법: java ConvertExample <입력파일경로> <출력파일경로>");
                System.out.println("예제: java ConvertExample test/빈파일/from.hwp test/빈파일/to.hwpx");
                System.exit(1);
            }
            
            String inputFilePath = args[0];
            String outputFilePath = args[1];
            
            // HWP 파일 읽기
            System.out.println("HWP 파일 읽는 중: " + inputFilePath);
            HWPFile fromFile = HWPReader.fromFile(inputFilePath);
            
            // HWPX로 변환
            System.out.println("HWPX로 변환 중...");
            HWPXFile toFile = Hwp2Hwpx.toHWPX(fromFile);
            
            // HWPX 파일 저장
            System.out.println("HWPX 파일 저장 중: " + outputFilePath);
            HWPXWriter.toFilepath(toFile, outputFilePath);
            
            System.out.println("변환 완료: " + outputFilePath);
            
        } catch (Exception e) {
            System.err.println("오류 발생: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
}


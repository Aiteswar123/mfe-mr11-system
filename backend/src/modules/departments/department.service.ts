import { PrismaClient, RoleCode } from '@prisma/client';
import { parseAndNormalizeWorkbook } from '../../utils/excel-normalizer';
import { executeMr11Pipeline } from '../mr11/mr11.engine';

export async function processAtomicWorkbookUpload(
  prisma: PrismaClient,
  deptCode: RoleCode,
  filePath: string,
  originalFilename: string,
  mimeType: string,
  fileSize: number,
  userId?: string
) {
  const parsedWorkbook = await parseAndNormalizeWorkbook(filePath);

  const dept = await prisma.department.findUnique({
    where: { code: deptCode },
  });

  if (!dept) {
    throw new Error(`Department with code ${deptCode} not found in database`);
  }

  // Validate userId to prevent foreign key errors
  let validUserId: string | null = null;
  if (userId) {
    const existingUser = await prisma.user.findUnique({ where: { id: userId } });
    if (existingUser) validUserId = existingUser.id;
  }
  if (!validUserId) {
    const adminUser = await prisma.user.findFirst({ where: { email: 'admin@mfeformwork.com' } });
    validUserId = adminUser ? adminUser.id : null;
  }

  const newVersion = await prisma.fileVersion.create({
    data: {
      departmentId: dept.id,
      originalFilename,
      storageKey: filePath,
      fileSize,
      mimeType,
      status: 'READY' as any,
      parsedWorkbook: parsedWorkbook as any,
      uploadedById: validUserId,
      uploadedAt: new Date(),
      processedAt: new Date(),
    },
  });

  await prisma.department.update({
    where: { id: dept.id },
    data: { activeVersionId: newVersion.id },
  });

  try {
    await prisma.auditLog.create({
      data: {
        userId: validUserId,
        action: 'UPLOAD_WORKBOOK',
        entityType: 'DEPARTMENT',
        entityId: dept.id,
        metadata: { filename: originalFilename, deptCode },
      },
    });
  } catch (err) {
    console.warn('Audit log write skipped:', err);
  }

  try {
    await executeMr11Pipeline(prisma);
  } catch (pipelineErr: any) {
    console.error('Auto MR11 pipeline calculation error:', pipelineErr);
  }

  return newVersion;
}